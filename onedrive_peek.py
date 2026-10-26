"""One-off, read-only: for each regional workbook compare what OneDrive's CLOUD holds (Microsoft Graph metadata) with the GCS copy."""
from urllib.parse import quote
import re
import sync   # signs in with the cached login (module level)

BASE = "SOM/Sell Through/Sell Through/Raw/Sell_Through_Offline_GT"
REG = re.compile(r"^\d+\.\s.*Sell Through\.xlsx$", re.I)
root = sync.get(f"{sync.GRAPH}/me/drive/root:/{quote(BASE)}:/children?$top=100").json()
print(f"{'file':34s} {'OneDrive cloud modified (UTC)':30s} {'GCS od_modified (UTC)':24s} {'same version?':14s} sizes OD/GCS")
for it in sorted(root.get("value", []), key=lambda x: x["name"]):
    if "file" not in it or not REG.match(it["name"]): continue
    b = sync.bucket.get_blob(f"{sync.GCS_PREFIX}/{it['name']}")
    md = (b.metadata or {}) if b else {}
    same = bool(b) and md.get("od_ctag") == it.get("cTag") and str(b.size) == str(it["size"])
    print(f"{it['name']:34s} {it['lastModifiedDateTime']:30s} {md.get('od_modified', '(none)'):24s} {'SAME' if same else 'GCS BEHIND?':14s} {it['size']:,} / {b.size if b else 0:,} | bucket updated {b.updated:%d %b %H:%M}Z" if b else f"{it['name']} not in GCS")
print("\nother files in the OneDrive folder (anything that is not one of the ten):")
for it in root.get("value", []):
    if not REG.match(it["name"]): print("  ", it["name"], it.get("size"), it["lastModifiedDateTime"])
