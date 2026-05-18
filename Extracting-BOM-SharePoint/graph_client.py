import json
import time
import pathlib
import requests

_CONFIG_PATH = pathlib.Path.home() / ".claude" / "msgraph_config.json"
_TOKEN_CACHE  = pathlib.Path.home() / ".claude" / "msgraph_token.json"

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"

DEFAULT_SCOPES = [
    "https://graph.microsoft.com/Files.Read.All",
    "https://graph.microsoft.com/Sites.Read.All",
    "offline_access",
]


def _load_credentials():
    if not _CONFIG_PATH.exists():
        raise FileNotFoundError(
            f"Credentials not found at {_CONFIG_PATH}. "
            "Run /msgraph-scaffold in Claude Code to set them up."
        )
    cfg = json.loads(_CONFIG_PATH.read_text())
    tenant_id = cfg.get("tenant_id", "").strip()
    client_id = cfg.get("client_id", "").strip()
    if not tenant_id or not client_id:
        raise ValueError(f"tenant_id and client_id must both be set in {_CONFIG_PATH}")
    return tenant_id, client_id


def _save_token(body):
    body = dict(body)
    body["expires_at"] = time.time() + int(body.get("expires_in", 3600))
    _TOKEN_CACHE.write_text(json.dumps(body))


def _read_cached_token():
    if not _TOKEN_CACHE.exists():
        return None, None
    try:
        cached = json.loads(_TOKEN_CACHE.read_text())
        at = cached.get("access_token")
        rt = cached.get("refresh_token")
        expires_at = float(cached.get("expires_at", 0))
        if at and expires_at - time.time() > 300:
            return at, rt
        return None, rt
    except Exception:
        return None, None


def _try_refresh(tenant_id, client_id):
    _, rt = _read_cached_token()
    if not rt:
        return None
    try:
        token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        resp = requests.post(token_url, data={
            "client_id":     client_id,
            "grant_type":    "refresh_token",
            "refresh_token": rt,
        })
        if resp.status_code != 200:
            return None
        body = resp.json()
        if "access_token" not in body:
            return None
        _save_token(body)
        return body["access_token"]
    except Exception:
        return None


def ensure_fresh_token():
    at, _ = _read_cached_token()
    if at:
        return at
    tenant_id, client_id = _load_credentials()
    new_token = _try_refresh(tenant_id, client_id)
    if new_token:
        try:
            import winsound
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        except Exception:
            pass
        print("  [auth] Token refreshed.", flush=True)
    return new_token


def acquire_token(scopes=None):
    tenant_id, client_id = _load_credentials()
    scopes = scopes or DEFAULT_SCOPES

    at, _ = _read_cached_token()
    if at:
        return at

    token = _try_refresh(tenant_id, client_id)
    if token:
        return token

    scope_str = " ".join(scopes)
    device_code_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/devicecode"
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    resp = requests.post(device_code_url, data={"client_id": client_id, "scope": scope_str})
    resp.raise_for_status()
    dc = resp.json()

    print(dc["message"], flush=True)

    interval = int(dc.get("interval", 5))
    deadline = time.time() + int(dc.get("expires_in", 900))

    while time.time() < deadline:
        time.sleep(interval)
        token_resp = requests.post(token_url, data={
            "client_id":   client_id,
            "grant_type":  "urn:ietf:params:oauth:grant-type:device_code",
            "device_code": dc["device_code"],
        })
        body = token_resp.json()
        error = body.get("error")
        if error == "authorization_pending":
            continue
        if error == "slow_down":
            interval += 5
            continue
        if error:
            raise RuntimeError(f"Auth error: {error} — {body.get('error_description')}")
        print("Signed in successfully.", flush=True)
        _save_token(body)
        return body["access_token"]

    raise TimeoutError("Device code flow timed out.")


def _retry_get(make_request, max_retries=6, headers=None):
    delay = 5
    token_refreshed = False
    for attempt in range(max_retries + 1):
        resp = make_request()
        if resp.status_code == 401 and headers is not None and not token_refreshed:
            tenant_id, client_id = _load_credentials()
            new_token = _try_refresh(tenant_id, client_id)
            if not new_token:
                time.sleep(2)
                new_token = _try_refresh(tenant_id, client_id)
            if new_token:
                headers["Authorization"] = f"Bearer {new_token}"
                token_refreshed = True
                continue
            resp.raise_for_status()
        if resp.status_code != 429:
            resp.raise_for_status()
            return resp
        retry_after = int(resp.headers.get("Retry-After", delay))
        wait = max(retry_after, delay)
        if attempt < max_retries:
            time.sleep(wait)
            delay = min(delay * 2, 120)
        else:
            resp.raise_for_status()
    return resp


def graph_get(access_token, url, params=None):
    headers = {"Authorization": f"Bearer {access_token}"}
    resp = _retry_get(lambda: requests.get(url, headers=headers, params=params), headers=headers)
    return resp.json()


def graph_get_all(access_token, url, params=None):
    """GET with automatic @odata.nextLink pagination. Returns flat list of all items."""
    items = []
    next_url = url
    next_params = params
    while next_url:
        body = graph_get(access_token, next_url, next_params)
        items.extend(body.get("value", []))
        next_url = body.get("@odata.nextLink")
        next_params = None
    return items
