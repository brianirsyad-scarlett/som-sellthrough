"""One-off: what does OneDrive's CLOUD hold for the Sell Through folder (name, size, last modified)? Read-only, metadata only."""
from urllib.parse import quote
import sync   # signs in with the same cached login as the sync (module level)

BASE = "SOM/Sell Through/Sell Through"
def show(label, path):
    try:
        it = sync.get(f"{sync.GRAPH}/me/drive/root:/{quote(path)}").json()
        print(f"{label:48s} cloud size {it.get('size', 0):>12,}  lastModified {it.get('lastModifiedDateTime')}")
    except Exception as e:
        print(f"{label:48s} NOT FOUND / error: {type(e).__name__} {str(e)[:90]}")

show("National Distribution Summary.xlsx", f"{BASE}/National Distribution Summary.xlsx")
for n in ("National Distribution Raw_20260701_20260930.xlsx", "National Distribution Raw_20261001_20261231.xlsx"):
    show("Raw/" + n, f"{BASE}/Raw/{n}")
show("Raw/Sell_Through_Offline_GT/National_Distribution.parquet", f"{BASE}/Raw/Sell_Through_Offline_GT/National_Distribution.parquet")
show("Raw/Sell_Through_Offline_GT/8. Sulawesi ...xlsx", f"{BASE}/Raw/Sell_Through_Offline_GT/8. Sulawesi Sell Through.xlsx")
print("\nfolder listing (what exists in the cloud folder):")
root = sync.get(f"{sync.GRAPH}/me/drive/root:/{quote(BASE)}:/children?$top=50").json()
for it in root.get("value", []):
    print("  ", "[dir] " if "folder" in it else "[file]", it["name"], "|", it.get("lastModifiedDateTime"))
