# SANS 10160-3 Wind Map

A single-page tool for reading the fundamental basic wind speed, v<sub>b,0</sub>, from SANS 10160-3 Figure 1 (3 s gust) for a site in South Africa.

Open [`WindMap.html`](WindMap.html) in a browser. It is self-contained: the map and municipality table are embedded, so it works offline.

## What it does

- Set the site by province and municipality, by GPS (degrees/minutes/seconds or decimal), or by clicking the map. With the map focused, arrow keys move the site by 1′ (Shift: 10′).
- v<sub>b,0</sub> comes from the municipality table when a municipality is selected. Otherwise it is read from the zone colour under the site. On a zone boundary the higher value is used and a warning is shown.
- A cross-check shows the zone under the site and the nearest municipality centroid with its table value, and flags when they disagree.
- Copy summary copies the coordinates and v<sub>b,0</sub> as text.

## Coordinates

South is entered as a positive number of degrees. The map is linear in degrees between its grid lines:

```
x = 114.5 + (E − 18) · 687.5 / 14     (pixels, image 958 wide)
y =  58.0 + (S − 22) · 588.5 / 12     (pixels, image 768 high)

Accepted range: E 16.4 – 33.0, S 22.0 – 34.9
```

`windmap2_coords.py` has the same conversion in Python, both ways.

## Files

| File | What it is |
|---|---|
| `WindMap.html` | The built app (standalone page) |
| `app/template.html` | App source: markup, styles and script |
| `app/build.py` | Embeds the map image and municipality table into the template |
| `WindMap2_Fig1_3s_gust.jpg` | SANS 10160-3 Figure 1, 3 s gust map |
| `WindMap1_Fig1_10min_avg.jpg` | Figure 1, 10 min average map (not used by the app) |
| `municipalities.csv` | Province, municipality, v<sub>b,0</sub> (3 s gust) and centroid South/East for 235 municipalities |
| `windmap2_coords.py` | Coordinate conversion in Python |
| `vba_UpdateMapGPS.bas` | Reference VBA for the map-plotting logic the app is based on |

## Building

After editing `app/template.html`:

```bash
python app/build.py
```

This writes `WindMap.html`.

## Source

SANS 10160-3, Basis of structural design and actions for buildings and industrial structures. Part 3: Wind actions.
