"""Port of the WindMap2 (SANS 10160-3 Fig 1, 3 s gust) coordinate system from
Design_Tool_SANS_10160_1.5.7.xlsb, VBA module modXLSMapCalcs.UpdateMapGPS (iPage=44).

Latitude is stored as positive degrees SOUTH, longitude as degrees EAST.
Positions are linear in degrees (plate carree) between two calibration anchors.
"""

# Calibration anchors, as fractions of the picture's width / height
X_ANC1, X_ANC2 = 20.0, 32.0        # deg E
Y_ANC1, Y_ANC2 = 22.0, 34.0        # deg S
X_ANC1_SCALE = 0.221               # 20 deg E sits at 22.1 % of picture width
X_SPAN_SCALE = 0.615               # 20->32 deg E spans 61.5 % of picture width
Y_ANC1_SCALE = 0.075               # 22 deg S sits at 7.5 % of picture height
Y_SPAN_SCALE = 0.765               # 22->34 deg S spans 76.5 % of picture height

# Allowed input window (outside this the VBA refuses to plot)
LON_MIN, LON_MAX = 16.4, 33.0
LAT_MIN, LAT_MAX = 22.0, 34.9

# Picture geometry: in Excel points on the sheet, and in pixels of the embedded JPG
SHEET_PICTURE = dict(left=329.1, top=24.6, width=448.8, height=381.0)
IMAGE_PX = dict(left=0, top=0, width=958, height=768)


def dms_to_deg(d, m, s):
    """Same as the sheet formula C8/C9: ((s/60)+m)/60+d."""
    return d + (m + s / 60.0) / 60.0


def deg_to_dms(deg):
    """Same as VBA CalcHMS(..., "HMS"); seconds rounded to 0.1 as the VBA does."""
    d = int(deg)
    m = int((deg - d) * 60)
    s = round(((deg - d) * 60 - m) * 60, 1)
    return d, m, s


def to_xy(south_deg, east_deg, frame=IMAGE_PX):
    """Map (deg S, deg E) to (x, y) in the given frame (pixels or sheet points)."""
    if not (LAT_MIN <= south_deg <= LAT_MAX and LON_MIN <= east_deg <= LON_MAX):
        raise ValueError("The requested position is not within the map boundaries")
    L, T, W, H = frame["left"], frame["top"], frame["width"], frame["height"]
    x1 = L + X_ANC1_SCALE * W
    y1 = T + Y_ANC1_SCALE * H
    x_per_deg = X_SPAN_SCALE * W / (X_ANC2 - X_ANC1)
    y_per_deg = Y_SPAN_SCALE * H / (Y_ANC2 - Y_ANC1)
    return x1 + (east_deg - X_ANC1) * x_per_deg, y1 + (south_deg - Y_ANC1) * y_per_deg


def to_latlon(x, y, frame=IMAGE_PX):
    """Inverse of to_xy: (x, y) in frame -> (deg S, deg E)."""
    L, T, W, H = frame["left"], frame["top"], frame["width"], frame["height"]
    x1 = L + X_ANC1_SCALE * W
    y1 = T + Y_ANC1_SCALE * H
    x_per_deg = X_SPAN_SCALE * W / (X_ANC2 - X_ANC1)
    y_per_deg = Y_SPAN_SCALE * H / (Y_ANC2 - Y_ANC1)
    return Y_ANC1 + (y - y1) / y_per_deg, X_ANC1 + (x - x1) / x_per_deg


if __name__ == "__main__":
    # Value saved in the workbook: 26 deg 0' 25.8" S, 28 deg 6' 8" E
    s, e = dms_to_deg(26, 0, 25.8), dms_to_deg(28, 6, 8)
    print(f"site: {s:.6f} S, {e:.6f} E")
    print("image px :", tuple(round(v, 1) for v in to_xy(s, e)))
    print("sheet pt :", tuple(round(v, 1) for v in to_xy(s, e, SHEET_PICTURE)))
