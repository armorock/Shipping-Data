"""
Parse Armorock precast part numbers into 15 attribute columns.
Implements the /item-name-detail-interpreter skill for all four naming generations
(Gen1 2014-2017, Gen2 2017-2023, Gen3 2019-2023, Gen4 current).

Returns a dict with keys:
  part_type, subcategory, generation, diameter, height, opening_diameter,
  troughing, wall_variant, section_suffix, lid_suffix, box_length, box_suffix,
  es, de, de_count
"""
import re

_BLANK = dict(
    part_type='', subcategory='', generation='',
    diameter='', height='', opening_diameter='',
    troughing='', wall_variant='', section_suffix='',
    lid_suffix='', box_length='', box_suffix='',
    es='', de='', de_count='',
)

_GEN4_DIAS = [192, 144, 120, 96, 84, 72, 60, 48]
_BOX_DIMS  = [154, 144, 120, 115, 104, 96, 91, 84, 79, 72, 65, 48, 36]

# Gen4 prefix order: longest first to prevent partial matches
_GEN4_PREFIXES = [
    'RMHLC', 'MHTLC', 'MHGRB',
    'RMHL', 'RMHC', 'MHTL', 'MHGR', 'MHCC', 'MHLC',
    'BOXS', 'BOXL', 'BOXF', 'BOXT',
    'RMH', 'MHB', 'MHC', 'MHL', 'MHS', 'MHT', 'BOX', 'MT',
]


def _base():
    return dict(_BLANK)


def _detect_generation(name):
    s = name.strip()
    if re.match(r'^60\d{3}-', s):
        s = '7' + s[1:]
    if re.match(r'^(MH|RMH|BOX|MT)', s, re.I):
        return 'Gen4', s
    if re.match(r'^[57]\d{4}-', s):
        return 'Gen2', s
    if re.match(r'^[SBCLANFTPMQV]\d{4}', s, re.I):
        return 'Gen3', s
    if re.match(r'^(5?GR\d|144|120|96|84|72|60|48|24RL|27RL|30RL|24TL|30TL)', s, re.I):
        return 'Gen1', s
    return 'Material', s


def _norm_tr(t):
    if not t:
        return ''
    t = str(t).strip()
    if re.match(r'^1\.3{1,2}$', t):
        return '.133'
    if re.match(r'^1(\.0?)?$', t):
        return '.1'
    if t.startswith('0.'):
        return t[1:]
    return t


def _conv_ht_gen1(s):
    return str(int(s) * 12) if len(s) == 1 else s


# ---------------------------------------------------------------------------
# Gen4 parsers
# ---------------------------------------------------------------------------

def _parse_mhb(r, rest):
    s = rest.upper()

    # Diameter
    dia = None
    for d in sorted(_GEN4_DIAS, reverse=True):
        ds = str(d)
        if s.startswith(ds):
            dia = ds
            s = s[len(ds):]
            break
    if not dia:
        return r
    r['diameter'] = dia

    # Height: 3 digits if the char after them is ., F, E, /, -, or end; else 2 digits
    if len(s) >= 3 and s[:3].isdigit():
        after3 = s[3:]
        if not after3 or after3[0] in '.FEf/-':
            r['height'] = s[:3]
            s = after3
        elif len(s) >= 2 and s[:2].isdigit():
            r['height'] = s[:2]
            s = s[2:]
    elif len(s) >= 2 and s[:2].isdigit():
        r['height'] = s[:2]
        s = s[2:]

    # Troughing — match longest candidates first
    for tr in ['.133', '.15', '.75', '.5', '.1', 'FF']:
        if s.upper().startswith(tr.upper()):
            r['troughing'] = tr
            s = s[len(tr):]
            break

    # Extended slab (.ES or ES)
    if s.upper().startswith('.ES'):
        r['es'] = 'Yes'
        s = s[3:]
    elif s.upper().startswith('ES'):
        r['es'] = 'Yes'
        s = s[2:]

    # Drop encasement
    if s.upper().startswith('/DE2') or s.upper().startswith('DE2'):
        r['de'] = 'Yes'
        r['de_count'] = '2'
        s = s[4:] if s[0] == '/' else s[3:]
    elif s.upper().startswith('/DE') or s.upper().startswith('DE'):
        r['de'] = 'Yes'
        r['de_count'] = '1'
        s = s[3:] if s[0] == '/' else s[2:]
    else:
        r['de_count'] = '0'

    # Wall variant
    m = re.match(r'^-([235])', s)
    if m:
        r['wall_variant'] = m.group(1)

    return r


def _parse_mhc(r, prefix, rest):
    for d in [72, 60, 48]:
        ds = str(d)
        if rest.startswith(ds):
            r['diameter'] = ds
            after = rest[len(ds):]
            m = re.match(r'^(\d{2})(-([235]))?', after)
            if m:
                r['opening_diameter'] = m.group(1)
                if m.group(3):
                    r['wall_variant'] = m.group(3)
            return r
    return r


def _parse_mhgr(r, rest):
    m = re.match(r'^(36|30|27|24)X?(\d)?', rest, re.I)
    if m:
        r['diameter'] = m.group(1)
        if m.group(2):
            r['height'] = m.group(2)
    return r


def _parse_mhgrb(r, rest):
    m = re.match(r'^(36|30|24)', rest)
    if m:
        r['diameter'] = m.group(1)
    return r


def _parse_mhl_family(r, prefix, rest):
    if prefix in ('MHL', 'MHLC'):
        dias = sorted(_GEN4_DIAS, reverse=True)
    else:  # MHTL, MHTLC
        dias = sorted([144, 120, 96, 84, 72, 60, 48, 36, 30], reverse=True)

    for d in dias:
        ds = str(d)
        if rest.startswith(ds):
            r['diameter'] = ds
            after = rest[len(ds):]
            # HATCH may appear before OD
            if re.match(r'^HATCH', after, re.I):
                r['lid_suffix'] = 'HATCH'
                return r
            m_od = re.match(r'^(\d{2})', after)
            if m_od:
                r['opening_diameter'] = m_od.group(1)
                after = after[2:]
                m_ls = re.match(r'^(CAST|HATCH|-TA)', after, re.I)
                if m_ls:
                    r['lid_suffix'] = m_ls.group(1).upper().lstrip('-')
                    after = after[len(m_ls.group(0)):]
                m_w = re.match(r'^-([235])', after)
                if m_w:
                    r['wall_variant'] = m_w.group(1)
            return r
    return r


def _parse_rmhl_family(r, prefix, rest):
    for d in [60, 48]:
        ds = str(d)
        if rest.startswith(ds):
            r['diameter'] = ds
            after = rest[len(ds):]
            m_od = re.match(r'^(\d{2})', after)
            if m_od:
                r['opening_diameter'] = m_od.group(1)
                after = after[2:]
                m_ls = re.match(r'^(CAST|HATCH)', after, re.I)
                if m_ls:
                    r['lid_suffix'] = m_ls.group(1).upper()
            return r
    return r


def _parse_mhs(r, rest):
    rest = re.sub(r'\s*\(inactive\)\s*$', '', rest, flags=re.I)
    for d in sorted(_GEN4_DIAS, reverse=True):
        ds = str(d)
        if rest.upper().startswith(ds):
            r['diameter'] = ds
            after = rest[len(ds):]
            m_ht = re.match(r'^(\d{2,3})', after)
            if m_ht:
                r['height'] = m_ht.group(1)
                after = after[len(m_ht.group(1)):]
                m_ss = re.match(r'^(NBH|NB|H|L)', after, re.I)
                if m_ss:
                    r['section_suffix'] = m_ss.group(1).upper()
                    after = after[len(m_ss.group(1)):]
                m_w = re.match(r'^-([235L])', after, re.I)
                if m_w:
                    r['wall_variant'] = m_w.group(1).upper()
            return r
    return r


def _parse_mht(r, rest):
    for d in [192, 144]:
        ds = str(d)
        if rest.startswith(ds):
            r['diameter'] = ds
            m = re.match(r'^-(\d+\.?\d*)', rest[len(ds):])
            if m:
                r['height'] = m.group(1)
            return r
    return r


def _parse_rmh(r, rest):
    for d in [60, 48]:
        ds = str(d)
        if rest.startswith(ds):
            r['diameter'] = ds
            after = rest[len(ds):]
            m_ht = re.match(r'^(\d{2})', after)
            if m_ht:
                r['height'] = m_ht.group(1)
                after = after[2:]
                if re.match(r'^H$', after, re.I):
                    r['section_suffix'] = 'H'
                elif re.match(r'^-?L$', after, re.I):
                    r['wall_variant'] = after
            return r
    return r


def _parse_rmhc(r, rest):
    m = re.match(r'^(60)(\d{2})', rest)
    if m:
        r['diameter'] = m.group(1)
        r['height'] = m.group(2)
    return r


def _parse_box(r, prefix, rest):
    type_label = {'BOX': 'Standard', 'BOXF': 'Flat Floor', 'BOXL': 'Lid',
                  'BOXS': 'Section', 'BOXT': 'Unknown'}.get(prefix, '')
    r['troughing'] = type_label

    s = rest.upper()
    if s.endswith('HATCH'):
        r['box_suffix'] = 'HATCH'
        s = s[:-5]
    elif s.endswith('FF'):
        r['box_suffix'] = 'FF'
        s = s[:-2]

    width = None
    for d in sorted(_BOX_DIMS, reverse=True):
        ds = str(d)
        if s.startswith(ds):
            width = ds
            s = s[len(ds):]
            break
    if width:
        r['diameter'] = width

    length = None
    for d in sorted(_BOX_DIMS, reverse=True):
        ds = str(d)
        if s.startswith(ds):
            length = ds
            s = s[len(ds):]
            break
    if length:
        r['box_length'] = length

    if s.isdigit():
        r['height'] = s

    return r


def _parse_gen4(name):
    s = re.sub(r'\s*\(inactive\)\s*$', '', name.strip(), flags=re.I).upper()
    r = _base()
    r['generation'] = 'Gen4'

    prefix = None
    for p in _GEN4_PREFIXES:
        if s.startswith(p):
            prefix = p
            rest = s[len(p):]
            break
    if not prefix:
        return r

    r['part_type'] = prefix
    if prefix.startswith('RMH'):
        r['subcategory'] = 'Rehab'
    elif prefix.startswith('BOX'):
        r['subcategory'] = 'Box'
    else:
        r['subcategory'] = 'Standard'

    if prefix == 'MHB':
        return _parse_mhb(r, rest)
    if prefix in ('MHC', 'MHCC'):
        return _parse_mhc(r, prefix, rest)
    if prefix == 'MHGR':
        return _parse_mhgr(r, rest)
    if prefix == 'MHGRB':
        return _parse_mhgrb(r, rest)
    if prefix in ('MHL', 'MHLC', 'MHTL', 'MHTLC'):
        return _parse_mhl_family(r, prefix, rest)
    if prefix in ('RMHL', 'RMHLC'):
        return _parse_rmhl_family(r, prefix, rest)
    if prefix == 'MHS':
        return _parse_mhs(r, rest)
    if prefix == 'MHT':
        return _parse_mht(r, rest)
    if prefix == 'RMH':
        return _parse_rmh(r, rest)
    if prefix == 'RMHC':
        return _parse_rmhc(r, rest)
    if prefix in ('BOX', 'BOXF', 'BOXL', 'BOXS', 'BOXT'):
        return _parse_box(r, prefix, rest)
    if prefix == 'MT':
        m = re.match(r'^(\d+)', rest)
        if m:
            r['diameter'] = m.group(1)
        return r
    return r


# ---------------------------------------------------------------------------
# Gen2 parser
# ---------------------------------------------------------------------------

_GEN2_SUFFIX = {
    'S':      ('MHS',   {}),
    'SNB':    ('MHS',   {'section_suffix': 'NB'}),
    'SNBH':   ('MHS',   {'section_suffix': 'NBH'}),
    'SR':     ('RMH',   {'subcategory': 'Rehab'}),
    'RS':     ('RMH',   {'subcategory': 'Rehab'}),
    'RL':     ('RMHL',  {'subcategory': 'Rehab'}),
    'RC':     ('RMHC',  {'subcategory': 'Rehab'}),
    'RCL':    ('RMHLC', {'subcategory': 'Rehab'}),
    'C':      ('MHC',   {}),
    'CC':     ('MHCC',  {}),
    'CL':     ('MHLC',  {}),
    'CLH':    ('MHLC',  {'lid_suffix': 'HATCH'}),
    'L':      ('MHL',   {}),
    'HL':     ('MHL',   {'lid_suffix': 'HATCH'}),
    'LH':     ('MHL',   {'lid_suffix': 'HATCH'}),
    'HLVP':   ('MHL',   {'lid_suffix': 'HATCH'}),
    'TL':     ('MHTL',  {}),
    'B':      ('MHB',   {}),
    'B1':     ('MHB',   {'troughing': '.1'}),
    'B50':    ('MHB',   {'troughing': '.5'}),
    'B75':    ('MHB',   {'troughing': '.75'}),
    'B100':   ('MHB',   {'troughing': '.1'}),
    'B116':   ('MHB',   {'troughing': '1.16'}),
    'B133':   ('MHB',   {'troughing': '.133'}),
    'BFF':    ('MHB',   {'troughing': 'FF'}),
    'B75ES':  ('MHB',   {'troughing': '.75', 'es': 'Yes'}),
    'B100ES': ('MHB',   {'troughing': '.1',  'es': 'Yes'}),
    'B133ES': ('MHB',   {'troughing': '.133','es': 'Yes'}),
    'BFFES':  ('MHB',   {'troughing': 'FF',  'es': 'Yes'}),
    'BXS':    ('BOXS',  {'subcategory': 'Box'}),
    'KBXS':   ('BOXS',  {'subcategory': 'Box'}),
    'BXSFF':  ('BOXS',  {'subcategory': 'Box', 'troughing': 'Flat Floor'}),
    'KBXSFF': ('BOXS',  {'subcategory': 'Box', 'troughing': 'Flat Floor'}),
    'TR':     ('MHB',   {'troughing': 'TR'}),
}

_GEN2_OD_TYPES = {'MHC', 'MHCC', 'MHLC', 'MHL', 'MHTL', 'RMHL', 'RMHLC', 'RMHC'}


def _parse_gen2(name):
    s = name.strip().upper()
    if re.match(r'^60\d{3}-', s):
        s = '7' + s[1:]

    r = _base()
    r['generation'] = 'Gen2'
    r['subcategory'] = 'Standard'

    m = re.match(r'^([57])(\d{2})(\d{2})-(.+)$', s)
    if not m:
        return r

    id_ft, ht_raw, suf = m.group(2), m.group(3), m.group(4)

    # Grade ring: ID is literal inches
    if suf == 'GR':
        r['part_type'] = 'MHGR'
        r['diameter'] = str(int(id_ft))
        r['height'] = str(int(ht_raw))
        return r

    dia = str(int(id_ft) * 12)

    # Patterns with embedded digits: S#C, SNB#D, SNB#C#D, Rs#C, etc.
    if re.match(r'^S\d+C', suf, re.I):
        r.update({'part_type': 'MHS', 'diameter': dia, 'height': str(int(ht_raw))})
        return r
    if re.match(r'^SNB', suf, re.I):
        r.update({'part_type': 'MHS', 'diameter': dia, 'height': str(int(ht_raw)),
                  'section_suffix': 'NB'})
        return r
    if re.match(r'^RS\d+C', suf, re.I) or re.match(r'^R?S\d', suf, re.I):
        r.update({'part_type': 'RMH', 'subcategory': 'Rehab',
                  'diameter': dia, 'height': str(int(ht_raw))})
        return r

    entry = _GEN2_SUFFIX.get(suf)
    if entry:
        pt, extras = entry
        r['part_type'] = pt
        r.update(extras)
        r['diameter'] = dia
        if pt in _GEN2_OD_TYPES:
            r['opening_diameter'] = str(int(ht_raw))
        else:
            r['height'] = str(int(ht_raw))
        return r

    # Unrecognized suffix — store what we know
    r['diameter'] = dia
    r['height'] = str(int(ht_raw))
    return r


# ---------------------------------------------------------------------------
# Gen3 parser
# ---------------------------------------------------------------------------

_GEN3_S_NB = {'0001', '1001', '1051', '0021', '0101', '0120'}
_GEN3_S_REHAB = {'2000', '2020', '2100', '2122', '2221', '2222'}

_GEN3_C_OPT = {
    '000': 'MHC', '050': 'MHC', '200': 'MHC', '500': 'MHC',
    '001': 'MHCC',
    '100': 'RMHC', '101': 'RMHC',
}

_GEN3_L_OPT = {
    '00000': ('MHL',  ''),    '00050': ('MHL',  ''),
    '00100': ('RMHL', ''),    '00051': ('RMHL', ''),
    '01000': ('MHL',  'HATCH'),
    '02000': ('MHLC', ''),    '02020': ('MHLC', ''), '02050': ('MHLC', ''),
    '02100': ('RMHLC', ''),
    '20000': ('MHL',  'CAST'), '20050': ('MHL',  'CAST'),
    '22000': ('MHLC', 'CAST'), '22050': ('MHLC', 'CAST'), '22222': ('MHLC', 'CAST'),
    '30000': ('MHL',  ''),    '30100': ('MHL',  ''), '40000': ('MHL', ''),
    '60000': ('MHL',  ''),
}


def _parse_gen3(name):
    s = re.sub(r'\s*\(inactive\)\s*$', '', name.strip(), flags=re.I).upper()
    r = _base()
    r['generation'] = 'Gen3'
    r['subcategory'] = 'Standard'

    m = re.match(r'^([SBCLANFTPMQV])(\d{2})(\d{2})-?(.*)$', s)
    if not m:
        return r

    letter, id_ft, ht_raw, opt = m.group(1), m.group(2), m.group(3), m.group(4)
    dia = str(int(id_ft) * 12)
    ht  = ht_raw
    r['diameter'] = dia

    if letter == 'A':
        r['part_type'] = 'MHGR'
        return r

    if letter == 'F':
        r['part_type'] = 'MHB'
        r['height'] = ht
        r['troughing'] = 'FF'
        return r

    if letter in ('B', 'T', 'M', 'V'):
        r['part_type'] = 'MHB'
        r['height'] = ht
        return r

    if letter == 'S':
        if opt in _GEN3_S_REHAB:
            r['part_type'] = 'RMH'
            r['subcategory'] = 'Rehab'
        else:
            r['part_type'] = 'MHS'
            if opt in _GEN3_S_NB:
                r['section_suffix'] = 'NB'
        r['height'] = ht
        return r

    if letter == 'P':
        r['part_type'] = 'MHS'
        r['height'] = ht
        return r

    if letter == 'N':
        r['part_type'] = 'MHL'
        r['height'] = ht
        return r

    if letter == 'C':
        pt = _GEN3_C_OPT.get(opt, 'MHC')
        r['part_type'] = pt
        if pt in ('RMHC',):
            r['subcategory'] = 'Rehab'
        r['opening_diameter'] = ht
        return r

    if letter == 'L':
        pt, ls = _GEN3_L_OPT.get(opt, ('MHL', ''))
        r['part_type'] = pt
        if pt in ('RMHL', 'RMHLC'):
            r['subcategory'] = 'Rehab'
        if ls:
            r['lid_suffix'] = ls
        r['opening_diameter'] = ht
        return r

    if letter == 'Q':
        r['part_type'] = 'MHS'
        r['height'] = ht
        return r

    return r


# ---------------------------------------------------------------------------
# Gen1 parser
# ---------------------------------------------------------------------------

def _parse_gen1(name):
    s = re.sub(r'\s*\(inactive\)\s*$', '', name.strip(), flags=re.I).upper()
    r = _base()
    r['generation'] = 'Gen1'
    r['subcategory'] = 'Standard'

    # Bare rehab lids
    m = re.match(r'^(24|27|30)RL$', s)
    if m:
        return {**r, 'part_type': 'RMHL', 'subcategory': 'Rehab',
                'diameter': '48', 'opening_diameter': m.group(1)}

    # Bare traffic lids
    m = re.match(r'^(24|30)TL$', s)
    if m:
        return {**r, 'part_type': 'MHTL', 'diameter': '48', 'opening_diameter': m.group(1)}

    # 5GR grade ring
    m = re.match(r'^5GR(\d{2})(\d{1,2})$', s)
    if m:
        return {**r, 'part_type': 'MHGR',
                'diameter': str(int(m.group(1))), 'height': str(int(m.group(2)))}

    # Standard GR grade ring
    m = re.match(r'^GR(\d{2})(\d{1,2})$', s)
    if m:
        return {**r, 'part_type': 'MHGR',
                'diameter': str(int(m.group(1))), 'height': str(int(m.group(2)))}

    # All other Gen1 names lead with a diameter
    m = re.match(r'^(144|120|96|84|72|60|48)(.*)', s)
    if not m:
        return r
    dia, rest = m.group(1), m.group(2)
    r['diameter'] = dia

    # No-type-letter base: one foot-digit then decimal troughing
    m2 = re.match(r'^(\d)(1?\.\d+)$', rest)
    if m2:
        return {**r, 'part_type': 'MHB',
                'height': str(int(m2.group(1)) * 12), 'troughing': _norm_tr(m2.group(2))}

    # Dash format: {ft}-{in}{type}
    m2 = re.match(r'^(\d)-(\d+)(.*)', rest)
    if m2:
        ht  = str(int(m2.group(1)) * 12 + int(m2.group(2)))
        tag = m2.group(3)
        if re.match(r'^BFF$', tag, re.I):
            return {**r, 'part_type': 'MHB', 'height': ht, 'troughing': 'FF'}
        m_bp = re.match(r'^BP?(\d*\.\d+)?$', tag, re.I)
        if m_bp:
            return {**r, 'part_type': 'MHB', 'height': ht, 'troughing': _norm_tr(m_bp.group(1) or '')}
        if re.match(r'^(SNB-DH|SNB|SB)$', tag, re.I):
            return {**r, 'part_type': 'MHS', 'height': ht, 'section_suffix': 'NB'}
        if re.match(r'^S$', tag, re.I):
            return {**r, 'part_type': 'MHS', 'height': ht}
        r['height'] = ht
        return r

    # BFF flat-floor base
    m2 = re.match(r'^(\d+)BFF$', rest, re.I)
    if m2:
        return {**r, 'part_type': 'MHB', 'height': _conv_ht_gen1(m2.group(1)), 'troughing': 'FF'}

    # B / BP + optional troughing
    m2 = re.match(r'^(\d+)BP?(\d*\.\d+)?$', rest, re.I)
    if m2:
        return {**r, 'part_type': 'MHB',
                'height': _conv_ht_gen1(m2.group(1)), 'troughing': _norm_tr(m2.group(2) or '')}

    # SR / RS rehab section
    m2 = re.match(r'^(\d+)(SR|RS)$', rest, re.I)
    if m2:
        return {**r, 'part_type': 'RMH', 'subcategory': 'Rehab',
                'height': _conv_ht_gen1(m2.group(1))}

    # SNB-DH / SNB / SB sections
    m2 = re.match(r'^(\d+)(SNB-DH|SNB|SB)$', rest, re.I)
    if m2:
        return {**r, 'part_type': 'MHS', 'height': _conv_ht_gen1(m2.group(1)), 'section_suffix': 'NB'}

    # Cored section s{n}C
    m2 = re.match(r'^(\d+)s\d+C$', rest, re.I)
    if m2:
        return {**r, 'part_type': 'MHS', 'height': _conv_ht_gen1(m2.group(1))}

    # Standard section
    m2 = re.match(r'^(\d+)S$', rest, re.I)
    if m2:
        return {**r, 'part_type': 'MHS', 'height': _conv_ht_gen1(m2.group(1))}

    # CRC / CR rehab cone
    m2 = re.match(r'^(\d{2})CRC$', rest, re.I)
    if m2:
        return {**r, 'part_type': 'RMHC', 'subcategory': 'Rehab', 'opening_diameter': m2.group(1)}

    m2 = re.match(r'^(\d{2})CR$', rest, re.I)
    if m2:
        return {**r, 'part_type': 'RMHC', 'subcategory': 'Rehab', 'opening_diameter': m2.group(1)}

    # LH hatch lid
    m2 = re.match(r'^(\d{2})?LH$', rest, re.I)
    if m2:
        return {**r, 'part_type': 'MHL', 'opening_diameter': m2.group(1) or '0', 'lid_suffix': 'HATCH'}

    # CC concentric cone
    m2 = re.match(r'^(\d{2})CC$', rest, re.I)
    if m2:
        return {**r, 'part_type': 'MHCC', 'opening_diameter': m2.group(1)}

    # C eccentric cone
    m2 = re.match(r'^(\d{2})C$', rest, re.I)
    if m2:
        return {**r, 'part_type': 'MHC', 'opening_diameter': m2.group(1)}

    # RL / LR rehab lid
    m2 = re.match(r'^(\d{2})(RL|LR)$', rest, re.I)
    if m2:
        return {**r, 'part_type': 'RMHL', 'subcategory': 'Rehab', 'opening_diameter': m2.group(1)}

    # Standard lid
    m2 = re.match(r'^(\d{2})L$', rest, re.I)
    if m2:
        return {**r, 'part_type': 'MHL', 'opening_diameter': m2.group(1)}

    # TL traffic lid
    m2 = re.match(r'^(\d{2})?TL$', rest, re.I)
    if m2:
        return {**r, 'part_type': 'MHTL', 'opening_diameter': m2.group(1) or '72'}

    return r


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

_PARSED_KEYS = list(_BLANK.keys())


def parse_part_name(pn):
    """Parse a part number string and return a dict with 15 attribute fields."""
    if not pn or not str(pn).strip():
        return _base()
    name = str(pn).strip()
    gen, normalized = _detect_generation(name)
    if gen == 'Gen4':
        return _parse_gen4(normalized)
    if gen == 'Gen2':
        return _parse_gen2(normalized)
    if gen == 'Gen3':
        return _parse_gen3(normalized)
    if gen == 'Gen1':
        return _parse_gen1(normalized)
    return _base()


def build_gen4_name(attrs):
    """Assemble a canonical Gen4 part name from a parsed-attribute dict.

    Input is the dict returned by parse_part_name() (keys without pn_ prefix), or any
    dict with the same keys.  Returns a Gen4 name string, or '#MISSING' when required
    fields are absent.
    """
    def v(k):
        return str(attrs.get(k) or '').strip()

    pt   = v('part_type')
    dia  = v('diameter')
    ht   = v('height')
    od   = v('opening_diameter')
    tr   = v('troughing')
    wall = v('wall_variant')
    ss   = v('section_suffix')
    ls   = v('lid_suffix')
    es   = v('es')
    de   = v('de_count')
    bl   = v('box_length')
    bs   = v('box_suffix')

    if not pt or not dia:
        return '#MISSING'

    def _od_pad(s):
        return '00' if s == '0' else s.zfill(2) if s else ''

    def _tr_sfx(t):
        if not t:
            return ''
        if t == 'FF':
            return 'FF'
        if t in ('1.0', '1'):
            return '.1'
        if t == '1.33':
            return '.133'
        return t if t.startswith('.') else ('.' + t.lstrip('0') if '.' in t else t)

    def _wall_sfx(w):
        return f'-{w}' if w else ''

    tr_s  = _tr_sfx(tr)
    es_s  = 'ES' if es == 'Yes' else ''
    de_s  = '/DE2' if de == '2' else '/DE' if de == '1' else ''
    w_s   = _wall_sfx(wall)
    od_s  = _od_pad(od)

    if pt == 'MHB':
        if not ht:
            return '#MISSING'
        return f'MHB{dia}{ht}{tr_s}{es_s}{de_s}{w_s}'

    if pt in ('MHC', 'MHCC'):
        if not od_s:
            return '#MISSING'
        return f'{pt}{dia}{od_s}{w_s}'

    if pt == 'MHGR':
        if not ht:
            return '#MISSING'
        return f'MHGR{dia}X{ht}'

    if pt == 'MHGRB':
        return f'MHGRB{dia}'

    if pt == 'MHL':
        if ls == 'HATCH':
            return f'MHL{dia}HATCH{w_s}'
        ls_s = 'CAST' if ls == 'CAST' else ('-TA' if ls == '-TA' else '')
        return f'MHL{dia}{od_s}{ls_s}{w_s}'

    if pt == 'MHLC':
        ls_s = 'CAST' if ls == 'CAST' else ''
        return f'MHLC{dia}{od_s}{ls_s}{w_s}'

    if pt == 'MHTL':
        return f'MHTL{dia}{od_s}{w_s}'

    if pt == 'MHTLC':
        return f'MHTLC{dia}{od_s}{w_s}'

    if pt == 'RMHL':
        ls_s = 'HATCH' if ls == 'HATCH' else ('CAST' if ls == 'CAST' else '')
        return f'RMHL{dia}{od_s}{ls_s}'

    if pt == 'RMHLC':
        return f'RMHLC{dia}{od_s}'

    if pt == 'MHS':
        if not ht:
            return '#MISSING'
        return f'MHS{dia}{ht}{ss}{w_s}'

    if pt == 'MHT':
        if not ht:
            return '#MISSING'
        return f'MHT{dia}-{ht}'

    if pt == 'RMH':
        if not ht:
            return '#MISSING'
        h_sfx = 'H' if ss == 'H' else ''
        return f'RMH{dia}{ht}{h_sfx}'

    if pt == 'RMHC':
        if not ht:
            return '#MISSING'
        return f'RMHC{dia}{ht}'

    if pt == 'MT':
        return f'MT{dia}'

    if pt in ('BOX', 'BOXF', 'BOXL', 'BOXS', 'BOXT'):
        _sub = {'Flat Floor': 'BOXF', 'Lid': 'BOXL', 'Section': 'BOXS'}
        prefix = _sub.get(tr, 'BOX')
        return f'{prefix}{dia}{bl}{ht}{bs}'

    return '#MISSING'
