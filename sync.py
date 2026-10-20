"""
OneDrive -> GCS sync for the ten regional Sell Through workbooks, run by GitHub Actions.

Source : OneDrive folder SOURCE_PATH (Microsoft Graph, delegated Files.Read)
Target : gs://GCS_BUCKET/GCS_PREFIX/   (the folder sell_through.py compiles)

A file is copied only when OneDrive holds something NEWER than the bucket. The laptop's
mirror daemon uploads every workbook the moment it is saved, usually before OneDrive has
synced it, so "different" is not enough: copying on a mere difference would overwrite the
fresh upload with the older OneDrive copy. Files deleted in OneDrive are kept in GCS.

The Microsoft refresh token rotates on every use, so the token cache is kept in the
(private) bucket between runs - NOT in the Actions cache, which a public repo can leak.
"""
import hashlib
import json
import logging
import os
import sys
import time
from datetime import datetime
from urllib.parse import quote

import msal
import requests
from google.cloud import storage

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("sync")



def _need(name: str) -> str:
    """A secret that exists but is blank (or pasted with quotes) otherwise fails deep inside MSAL."""
    v = os.environ.get(name, "").strip().strip("\"'").strip()
    if not v:
        sys.exit(f"{name} is empty - re-set that repository secret (the bare value, no quotes)")
    return v


TENANT_ID = _need("AZURE_TENANT_ID")
CLIENT_ID = _need("AZURE_CLIENT_ID")
SEED = os.environ.get("MS_TOKEN_CACHE", "").strip()   # first run / after a re-login
SOURCE_PATH = os.environ.get(
    "SOURCE_PATH", "SOM/Sell Through/Sell Through/Raw/Sell_Through_Offline_GT"
).strip("/")
GCS_BUCKET = os.environ.get("GCS_BUCKET", "bucket_som")
GCS_PREFIX = os.environ.get(
    "GCS_PREFIX", "sales_parquet/raw/sell_through/Sell_Through_Offline_GT"
).strip("/")
STATE_BLOB = os.environ.get("STATE_BLOB", "_state/som-sellthrough/msal_token_cache.json")

import re

# Only the regional workbooks ("1. Sumbagut Sell Through.xlsx" ... "8. Sulawesi Sell Through.xlsx") are sales data.
# Anything else dropped into that OneDrive folder (an Excel "AutoRecovered" copy, a scratch list ...) used to be compiled too,
# added its own columns to the table and made the publish check refuse every build.
REGIONAL_RE = re.compile(r"^\d+\.\s.*Sell Through\.xlsx$", re.IGNORECASE)


def is_regional(name: str) -> bool:
    return bool(REGIONAL_RE.match(name.rsplit("/", 1)[-1].rsplit("\\", 1)[-1].strip()))

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPE = ["Files.Read"]
SEED_SHA = hashlib.sha256(SEED.encode()).hexdigest() if SEED else ""

bucket = storage.Client().bucket(GCS_BUCKET)
cache = msal.SerializableTokenCache()


def load_cache() -> None:
    blob = bucket.blob(STATE_BLOB)
    state = json.loads(blob.download_as_text()) if blob.exists() else None
    if state and state.get("seed_sha256") == SEED_SHA:
        cache.deserialize(state["cache"])
        log.info("token cache: loaded from gs://%s/%s", GCS_BUCKET, STATE_BLOB)
    elif SEED:
        cache.deserialize(SEED)
        log.info("token cache: seeded from the MS_TOKEN_CACHE secret")
    else:
        sys.exit("No token cache: set the MS_TOKEN_CACHE secret (token_cache.json from login.py)")


def persist_cache() -> None:
    if cache.has_state_changed:
        bucket.blob(STATE_BLOB).upload_from_string(
            json.dumps({"seed_sha256": SEED_SHA, "cache": cache.serialize()}),
            content_type="application/json",
        )
        log.info("token cache: saved")


load_cache()
app = msal.PublicClientApplication(
    CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}", token_cache=cache
)


def token() -> str:
    accounts = app.get_accounts()
    if not accounts:
        sys.exit("No account in the token cache - sign in again with login.py and update MS_TOKEN_CACHE")
    res = app.acquire_token_silent(SCOPE, account=accounts[0])
    if not res or "access_token" not in res:
        sys.exit(f"Token refresh failed - sign in again with login.py and update MS_TOKEN_CACHE: {res}")
    persist_cache()   # a rotated refresh token must survive even if the run dies later
    return res["access_token"]


def get(url: str) -> requests.Response:
    for attempt in range(6):
        r = requests.get(url, headers={"Authorization": f"Bearer {token()}"}, timeout=60)
        if r.status_code in (429, 500, 502, 503, 504):
            wait = int(r.headers.get("Retry-After", 2 ** attempt))
            log.warning("Graph %s, retrying in %ss", r.status_code, wait)
            time.sleep(wait)
            continue
        r.raise_for_status()
        return r
    r.raise_for_status()
    return r


def walk(item_id: str, rel: str = ""):
    url = f"{GRAPH}/me/drive/items/{item_id}/children?$top=200"
    while url:
        data = get(url).json()
        for it in data["value"]:
            path = f"{rel}/{it['name']}" if rel else it["name"]
            if "folder" in it:
                yield from walk(it["id"], path)
            elif "file" in it:
                yield path, it
        url = data.get("@odata.nextLink")


def source_files():
    root = get(f"{GRAPH}/me/drive/root:/{quote(SOURCE_PATH)}:").json()
    if "file" in root:
        yield root["name"], root
    else:
        yield from walk(root["id"])


def wants(rel: str) -> bool:
    name = rel.rsplit("/", 1)[-1]
    return name.lower().endswith(".xlsx") and not name.startswith("~$") and is_regional(name)   # skip Excel lock files + non-regional files


def needs_copy(blob, item) -> tuple[bool, str]:
    if blob is None:
        return True, "not in bucket"
    meta = blob.metadata or {}
    if meta.get("od_ctag") == item.get("cTag") and str(blob.size) == str(item["size"]):
        return False, "unchanged"
    od_modified = datetime.fromisoformat(item["lastModifiedDateTime"].replace("Z", "+00:00"))
    if blob.updated >= od_modified:
        return False, f"bucket copy is newer ({blob.updated:%m-%d %H:%M:%S}Z >= OneDrive {od_modified:%m-%d %H:%M:%S}Z)"
    return True, f"OneDrive is newer ({od_modified:%m-%d %H:%M:%S}Z > bucket {blob.updated:%m-%d %H:%M:%S}Z)"


def main() -> int:
    uploaded, skipped, errors = [], 0, []
    for rel, item in source_files():
        if not wants(rel):
            n = rel.rsplit("/", 1)[-1]
            if n.lower().endswith(".xlsx") and not n.startswith("~$"):
                log.warning("ignoring %s: not a regional 'N. <Region> Sell Through.xlsx' workbook", rel)
            continue
        key = f"{GCS_PREFIX}/{rel}"
        try:
            copy, why = needs_copy(bucket.get_blob(key), item)
            if not copy:
                log.info("skip   %s - %s", rel, why)
                skipped += 1
                continue
            dl = item.get("@microsoft.graph.downloadUrl") or get(
                f"{GRAPH}/me/drive/items/{item['id']}"
            ).json()["@microsoft.graph.downloadUrl"]
            new = bucket.blob(key)
            new.metadata = {
                "od_ctag": item.get("cTag"),
                "od_id": item["id"],
                "od_modified": item["lastModifiedDateTime"],
            }
            with requests.get(dl, stream=True, timeout=(10, 300)) as r:   # pre-authed URL: no bearer
                r.raise_for_status()
                r.raw.decode_content = True
                new.upload_from_file(r.raw, content_type=item["file"].get("mimeType"))
            log.info("copied %s (%s bytes) - %s", rel, item["size"], why)
            uploaded.append(rel)
        except Exception:
            log.exception("failed %s", rel)
            errors.append(rel)

    persist_cache()
    log.info("done: %d copied, %d skipped, %d errors", len(uploaded), skipped, len(errors))
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
