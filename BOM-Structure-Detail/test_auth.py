"""
Quick auth test — verifies client credentials flow works before running a full extraction.
Run: python test_auth.py
"""
from graph_client import _load_credentials, _try_client_credentials, graph_get

tenant_id, client_id, client_secret = _load_credentials()

print("Testing client credentials flow...")
token = _try_client_credentials(tenant_id, client_id, client_secret)

if not token:
    print("FAILED — could not get token. Check client_secret in msgraph_config.json.")
    raise SystemExit(1)

print("Token acquired.")

print("Testing Graph API call (listing SharePoint sites)...")
result = graph_get(token, "https://graph.microsoft.com/v1.0/sites?search=jobdata2026")
sites = result.get("value", [])

if sites:
    print(f"SUCCESS — {len(sites)} site(s) found:")
    for s in sites:
        print(f"  {s.get('displayName')} — {s.get('webUrl')}")
else:
    print("WARNING — token worked but no sites returned. Check API permissions in Azure AD.")
