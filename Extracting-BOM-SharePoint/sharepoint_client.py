from urllib.parse import quote
import requests
from graph_client import GRAPH_ROOT, graph_get, graph_get_all, _retry_get

SKIP_FOLDERS = {"forms", "plugin_data", "robotinterface", "__macosx", "antiquated files"}


def get_site(access_token, hostname, site_path):
    encoded = quote(site_path, safe="/")
    url = f"{GRAPH_ROOT}/sites/{hostname}:{encoded}"
    return graph_get(access_token, url)


def get_drive(access_token, site_id, drive_name="Documents"):
    drives = graph_get_all(access_token, f"{GRAPH_ROOT}/sites/{site_id}/drives")
    for d in drives:
        if d.get("name", "").lower() == drive_name.lower():
            return d
    available = [d.get("name") for d in drives]
    raise ValueError(f"Drive '{drive_name}' not found. Available: {available}")


def list_children(access_token, drive_id, folder_id="root"):
    """Return direct children of a folder (non-recursive)."""
    url = f"{GRAPH_ROOT}/drives/{drive_id}/items/{folder_id}/children"
    params = {"$select": "id,name,size,webUrl,file,folder,parentReference"}
    return graph_get_all(access_token, url, params)


def iter_files(access_token, drive_id, folder_id="root", recursive=True):
    """Yield all file DriveItems under a folder."""
    for item in list_children(access_token, drive_id, folder_id):
        if "file" in item:
            yield item
        elif "folder" in item and recursive:
            if item["name"].lower() not in SKIP_FOLDERS:
                yield from iter_files(access_token, drive_id, item["id"], recursive=True)


def download_file(access_token, drive_id, item_id):
    """Download a file and return its bytes."""
    headers = {"Authorization": f"Bearer {access_token}"}
    meta_url = f"{GRAPH_ROOT}/drives/{drive_id}/items/{item_id}"
    meta = _retry_get(lambda: requests.get(meta_url, headers=headers))
    dl_url = meta.json().get("@microsoft.graph.downloadUrl")
    resp = _retry_get(lambda: requests.get(dl_url))
    return resp.content
