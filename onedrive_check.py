"""Read-only: what does the OneDrive CLOUD hold right now? Downloads the ten regional workbooks straight from
Microsoft Graph (not from GCS), scans every visible sheet, and publishes
  gs://bucket_som/_reports/som-sellthrough/onedrive_last_update.csv   (last data date per distributor)
  gs://bucket_som/_reports/som-sellthrough/onedrive_cloud/<file>.xlsx  (the cloud copies, for deeper checks)
Run on demand (workflow 'OneDrive check'); it never touches the pipeline folder."""
import datetime as dt
import io
import re
from urllib.parse import quote

import pandas as pd
from openpyxl import load_workbook
import sync   # signs in with the cached login

BASE = sync.SOURCE_PATH
OUT = "_reports/som-sellthrough"
TODAY = dt.date.today()


def to_date(v):
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 36526 <= v <= 60000:
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))).date()
    if isinstance(v, str):
        m = re.fullmatch(r"\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*", v)
        if m:
            try: return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except ValueError: return None
    return None


rows = []
root = sync.get(f"{sync.GRAPH}/me/drive/root:/{quote(BASE)}:/children?$top=100").json()
for it in sorted(root.get("value", []), key=lambda x: x["name"]):
    if "file" not in it or not sync.is_regional(it["name"]): continue
    data = sync.get(f"{sync.GRAPH}/me/drive/items/{it['id']}/content").content
    sync.bucket.blob(f"{OUT}/onedrive_cloud/{it['name']}").upload_from_string(data)
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    area = re.sub(r"\s*Sell Through\.xlsx$", "", it["name"], flags=re.I).upper()
    for ws in wb.worksheets:
        if ws.sheet_state != "visible" or ws.title.strip().lower() == "master_data": continue
        itr = ws.iter_rows(values_only=True); hdr = [str(x).strip() if x is not None else "" for x in next(itr)]
        di = hdr.index("Date"); n = 0; last = None; undated = 0
        for r in itr:
            if not any(v is not None for v in r): continue
            n += 1; d = to_date(r[di])
            if d is None: undated += 1
            elif d <= TODAY and (last is None or d > last): last = d
        rows.append({"Area": area, "Distributor": ws.title, "OneDrive last date": last, "rows": n, "undated": undated,
                     "OneDrive modified (UTC)": it["lastModifiedDateTime"], "size": it["size"]})
    wb.close()
    print("scanned", it["name"], flush=True)
df = pd.DataFrame(rows)
sync.bucket.blob(f"{OUT}/onedrive_last_update.csv").upload_from_string(df.to_csv(index=False), content_type="text/csv")
print(df.to_string())
