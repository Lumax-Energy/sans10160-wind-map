"""Build WindMap.html from template.html: embeds the 3 s gust map and the municipality table."""
import base64
import csv
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent


def data_uri(path):
    return "data:image/jpeg;base64," + base64.b64encode(path.read_bytes()).decode("ascii")


muns = []
with open(ROOT / "municipalities.csv", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        name = r["Municipality"].replace("�", "â")  # workbook stores "Khâi-Ma" with a broken character
        muns.append([int(r["ProvinceID"]), int(r["MunicipalityID"]), name,
                     int(float(r["vb0_3s_gust_m_s"])), round(float(r["South_deg"]), 7), round(float(r["East_deg"]), 7)])

html = (HERE / "template.html").read_text(encoding="utf-8")
html = html.replace("__IMG_G__", data_uri(ROOT / "WindMap2_Fig1_3s_gust.jpg"))
html = html.replace("__MUNS__", json.dumps(muns, ensure_ascii=False, separators=(",", ":")))
# Artifact body (the publisher adds doctype/head) and a standalone page for opening locally
body = ROOT / "app" / "WindMap.artifact.html"
body.write_text(html, encoding="utf-8")
out = ROOT / "WindMap.html"
out.write_text('<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n'
               '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
               "</head>\n<body>\n" + html + "\n</body>\n</html>\n", encoding="utf-8")
print(f"{out} ({out.stat().st_size / 1024:.0f} KB, {len(muns)} municipalities)")
