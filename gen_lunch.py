from PIL import Image, ImageDraw, ImageFont

# ── FONTS ────────────────────────────────────────────────────────────────────
B = "/usr/share/fonts/truetype/liberation/"
F_REG  = ImageFont.truetype(B + "LiberationSans-Regular.ttf",  21)
F_BOL  = ImageFont.truetype(B + "LiberationSans-Bold.ttf",     21)
F_SM   = ImageFont.truetype(B + "LiberationSans-Regular.ttf",  17)
F_SMB  = ImageFont.truetype(B + "LiberationSans-Bold.ttf",     17)
F_XS   = ImageFont.truetype(B + "LiberationSans-Regular.ttf",  14)
F_XSB  = ImageFont.truetype(B + "LiberationSans-Bold.ttf",     14)
F_TIT  = ImageFont.truetype(B + "LiberationSans-Bold.ttf",     26)

# ── COLORS (dark theme matching screenshot) ───────────────────────────────────
C = {
    "bg":        "#1A1A1A",
    "hdr":       "#2A1507",   # brand charcoal
    "hdr_txt":   "#FEDA68",   # brand gold
    "col_hdr":   "#242424",
    "col_txt":   "#FFFFFF",
    "col_sub":   "#9A9590",
    "row1":      "#1A1A1A",
    "row2":      "#222222",
    "lunch_bg":  "#2A1F10",
    "lunch_txt": "#FEDA68",
    "time_txt":  "#FFFFFF",
    "va_txt":    "#FFFFFF",
    "leads_txt": "#8A8580",
    "moved_bg":  "#3D2800",
    "moved_txt": "#F5A623",
    "grid":      "#333333",
    "warn_bg":   "#2D1A00",
    "warn_txt":  "#F5A623",
    "pst_txt":   "#B8A0D8",
    "maykee":    "#4DBFBF",
    "reyn":      "#7EC86E",
}

# ── TODAY'S DATA (2026-05-22, Friday — Ann OFF) ───────────────────────────────
# Redistribution by lead volume:
#   Maykee base 1071 | Reyn base 812
#   Ed(232)  → Reyn  (812→1044)
#   Mark(236)→ Reyn  (1044→1280)
#   Myra(167)→ Maykee(1071→1238)

DATE  = "Friday, May 22, 2026"
NOTE  = "⚠  Ann is OFF today  —  Myra → Maykee   |   Mark + Ed → Reyn"

# Columns: (label, sub-label, color)
COLS = [
    ("MAYKEE",
     "Kirk · Jayson · Lou · Jell · Rovy · Jan  +  Myra*",
     C["maykee"]),
    ("REYN",
     "Bea · Klaie · Ena · Rincee · Dan · Veronica  +  Mark* · Ed*",
     C["reyn"]),
]

# Rows: (time_label, [cell_per_col])
# cell = (lines[], leads_total, is_lunch, has_moved)
# leads -1 = lunch cell, 0 = empty
ROWS = [
    ("11:00\n– 12:00", [
        (["Rovy / Jayson"],   241, False, False),
        (["Bea"],             207, False, False),
    ]),
    ("12:00\n– 1:00", [
        (["\U0001f37d  MAYKEE LUNCH"],  0, True, False),
        (["Mark* / Rincee"],  315, False, True),
    ]),
    ("1:00\n– 2:00", [
        (["Lou / Kirk"],      136, False, False),
        (["\U0001f37d  REYN LUNCH"],    0, True, False),
    ]),
    ("2:00\n– 3:00", [
        (["Jell"],            348, False, False),
        (["Klaie / Veronica"],258, False, False),
    ]),
    ("3:00\n– 4:00", [
        (["Jan / Myra*"],     376, False, True),
        (["Ena / Dan"],       197, False, False),
    ]),
    ("4:00 – 5:00\nPST only", [
        (["Jessica"],         137, False, False),
        (["Ed* / Neil"],      303, False, True),
    ]),
]

# ── LAYOUT ───────────────────────────────────────────────────────────────────
W        = 900
COL_T    = 145   # time column width
COL_W    = (W - COL_T - 2) // len(COLS)   # flex VA column width
HDR_H    = 86
WARN_H   = 46
COLHDR_H = 72
ROW_H    = 72
FOOT_H   = 38
H = HDR_H + WARN_H + COLHDR_H + len(ROWS) * ROW_H + FOOT_H

img = Image.new("RGB", (W, H), C["bg"])
d   = ImageDraw.Draw(img)

def txt_c(draw, cx, cy, text, font, fill):
    """Draw text centered at (cx, cy)."""
    bb = draw.textbbox((0, 0), text, font=font)
    tw, th = bb[2] - bb[0], bb[3] - bb[1]
    draw.text((cx - tw // 2, cy - th // 2), text, font=font, fill=fill)

def hline(y, color=None):
    d.line([(0, y), (W, y)], fill=color or C["grid"], width=1)

def vline(x, y0, y1, color=None):
    d.line([(x, y0), (x, y1)], fill=color or C["grid"], width=1)

y = 0

# ── HEADER ────────────────────────────────────────────────────────────────────
d.rectangle([0, 0, W, HDR_H], fill=C["hdr"])
d.text((22, 16), "🍽  LUNCH SCHEDULE", font=F_TIT, fill=C["hdr_txt"])
d.text((22, 52), DATE, font=F_SM, fill="#BFB9AC")
d.text((W - 22 - d.textbbox((0,0),"EST window: 11 AM – 4 PM  ·  PST: 4–5 PM",font=F_XS)[2],
        58), "EST window: 11 AM – 4 PM  ·  PST: 4–5 PM", font=F_XS, fill="#6A6560")
y = HDR_H

# ── WARNING BANNER ────────────────────────────────────────────────────────────
d.rectangle([0, y, W, y + WARN_H], fill=C["warn_bg"])
nw = d.textbbox((0,0), NOTE, font=F_SMB)[2]
d.text(((W - nw) // 2, y + 13), NOTE, font=F_SMB, fill=C["warn_txt"])
hline(y + WARN_H, C["warn_txt"] + "55")
y += WARN_H

# ── COLUMN HEADERS ────────────────────────────────────────────────────────────
d.rectangle([0, y, W, y + COLHDR_H], fill=C["col_hdr"])
# Time header
d.text((22, y + 22), "TIME (EST)", font=F_XSB, fill=C["col_sub"])

for i, (name, sub, color) in enumerate(COLS):
    cx = COL_T + i * COL_W + COL_W // 2
    txt_c(d, cx, y + 22, name, F_BOL, color)
    # sub label (wrap if too long)
    sw = d.textbbox((0,0), sub, font=F_XS)[2]
    if sw > COL_W - 12:
        # split at the + sign
        parts = sub.split("  +  ")
        d.text((COL_T + i * COL_W + 10, y + 40), parts[0].strip(), font=F_XS, fill=C["col_sub"])
        if len(parts) > 1:
            d.text((COL_T + i * COL_W + 10, y + 56), "+ " + parts[1].strip(), font=F_XS, fill=C["moved_txt"])
    else:
        sw2 = d.textbbox((0,0), sub, font=F_XS)[2]
        d.text((COL_T + i * COL_W + (COL_W - sw2) // 2, y + 46), sub, font=F_XS, fill=C["col_sub"])

hline(y + COLHDR_H, "#444444")
# vertical dividers in col header
vline(COL_T, y, y + COLHDR_H, "#333333")
for i in range(1, len(COLS)):
    vline(COL_T + i * COL_W, y, y + COLHDR_H, "#333333")
y += COLHDR_H

# ── ROWS ──────────────────────────────────────────────────────────────────────
for ri, (time_lbl, cells) in enumerate(ROWS):
    row_bg = C["row1"] if ri % 2 == 0 else C["row2"]
    d.rectangle([0, y, W, y + ROW_H], fill=row_bg)

    # Time label
    lines = time_lbl.split("\n")
    if len(lines) == 2:
        d.text((10, y + 15), lines[0], font=F_BOL, fill=C["time_txt"])
        d.text((10, y + 38), lines[1], font=F_SM,  fill=C["col_sub"])
    else:
        txt_c(d, COL_T // 2, y + ROW_H // 2, lines[0], F_BOL, C["time_txt"])

    vline(COL_T, y, y + ROW_H)

    for ci, (lines, leads, is_lunch, has_moved) in enumerate(cells):
        cx0 = COL_T + ci * COL_W
        cx1 = cx0 + COL_W
        mid_x = (cx0 + cx1) // 2

        if is_lunch:
            d.rectangle([cx0 + 2, y + 2, cx1 - 2, y + ROW_H - 2], fill=C["lunch_bg"])
            txt_c(d, mid_x, y + ROW_H // 2, lines[0], F_SMB, C["lunch_txt"])
        else:
            va_txt = lines[0]
            # VA name line
            color = COLS[ci][2]
            if has_moved:
                txt_c(d, mid_x, y + 20, va_txt, F_BOL, C["moved_txt"])
            else:
                txt_c(d, mid_x, y + 20, va_txt, F_BOL, C["va_txt"])

            # Leads line
            if leads > 0:
                leads_str = f"{leads:,} leads"
                txt_c(d, mid_x, y + 47, leads_str, F_XS, C["leads_txt"])

        if ci < len(COLS) - 1:
            vline(cx1, y, y + ROW_H)

    hline(y + ROW_H)
    y += ROW_H

# ── FOOTER ────────────────────────────────────────────────────────────────────
d.rectangle([0, y, W, H], fill=C["hdr"])
note = "* moved from Ann's group  ·  Leads = MTD total  ·  Roof Ignite"
nw = d.textbbox((0,0), note, font=F_XS)[2]
d.text(((W - nw) // 2, y + 12), note, font=F_XS, fill="#6A6560")

out = "/home/user/va-calendar/lunch_schedule_2026-05-22.png"
img.save(out, "PNG")
print(f"Saved: {out}  ({W}×{H}px)")
