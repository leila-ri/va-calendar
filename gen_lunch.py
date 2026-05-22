from PIL import Image, ImageDraw, ImageFont

# ── TODAY: 2026-05-22 (Friday) ──────────────────────────────────────────────
# Ann is OFF today (flex schedule).
# Redistribution: balance Ed/Mark/Myra to Maykee & Reyn by lead volume.
#   Maykee base: 1071  Reyn base: 812
#   Ed(232)   → Reyn  (812→1044)
#   Mark(236) → Reyn  (1044→1280)
#   Myra(167) → Maykee(1071→1238)

DATE = "Friday, May 22, 2026"
ANN_OFF = True

# (name, lunch_time, covered_by, tz, moved)
ROWS = [
    # ── MAYKEE group ─────────────────────────────────────────────────────────
    ("Jell",     "11:00 AM", "Maykee", "EST", False),
    ("Jan",      "11:30 AM", "Maykee", "EST", False),
    ("Rovy",     "12:00 PM", "Maykee", "EST", False),
    ("Myra",     "12:30 PM", "Maykee", "EST", True),   # moved from Ann
    ("Lou",      "1:00 PM",  "Maykee", "EST", False),
    ("Kirk",     "1:30 PM",  "Maykee", "EST", False),
    ("Jayson",   "2:00 PM",  "Maykee", "EST", False),
    ("Jessica",  "4:00 PM",  "Maykee", "PST", False),
    # ── REYN group ───────────────────────────────────────────────────────────
    ("Mark",     "11:00 AM", "Reyn",   "EST", True),   # moved from Ann
    ("Bea",      "11:30 AM", "Reyn",   "EST", False),
    ("Ena",      "12:00 PM", "Reyn",   "EST", False),
    ("Klaie",    "12:30 PM", "Reyn",   "EST", False),
    ("Veronica", "1:00 PM",  "Reyn",   "EST", False),
    ("Rincee",   "1:30 PM",  "Reyn",   "EST", False),
    ("Dan",      "2:00 PM",  "Reyn",   "EST", False),
    ("Ed",       "4:00 PM",  "Reyn",   "PST", True),   # moved from Ann
    ("Neil",     "4:00 PM",  "Reyn",   "PST", False),
]

# ── FONTS ────────────────────────────────────────────────────────────────────
BASE  = "/usr/share/fonts/truetype/liberation/"
F_REG = ImageFont.truetype(BASE + "LiberationSans-Regular.ttf",  22)
F_BOL = ImageFont.truetype(BASE + "LiberationSans-Bold.ttf",     22)
F_SM  = ImageFont.truetype(BASE + "LiberationSans-Regular.ttf",  18)
F_SMB = ImageFont.truetype(BASE + "LiberationSans-Bold.ttf",     18)
F_XS  = ImageFont.truetype(BASE + "LiberationSans-Regular.ttf",  15)
F_TIT = ImageFont.truetype(BASE + "LiberationSans-Bold.ttf",     28)

# ── COLORS ───────────────────────────────────────────────────────────────────
BG        = "#F7F6F2"
HEADER_BG = "#2A1507"
HEADER_FG = "#FEDA68"
ROW_ALT   = "#F0EDE8"
ROW_NORM  = "#FFFFFF"
ROW_MOVED = "#FFF8E8"
BORDER    = "#D4D0C8"
EST_BG    = "#EDF4FA";  EST_FG    = "#2E6B9B"
PST_BG    = "#F3F0FA";  PST_FG    = "#6B4FA0"
MOVED_BG  = "#FEF5EC";  MOVED_FG  = "#C06D1F"
MAYKEE_BG = "#ECF6F6";  MAYKEE_FG = "#2E7D7E"
REYN_BG   = "#F0F6EA";  REYN_FG   = "#4A7C2E"
WARN_BG   = "#FEF5EC";  WARN_FG   = "#C06D1F"
SUBHD_BG  = "#EEE8E0"
MU        = "#7A7567"
DARK      = "#1C1C1C"

# ── DIMENSIONS ───────────────────────────────────────────────────────────────
W        = 880
PAD      = 24
ROW_H    = 40
COL_W    = [200, 160, 140, 90]  # VA | Lunch Time | Covered By | TZ
COLS     = ["VA", "LUNCH TIME (EST)", "COVERED BY", "TZ"]

total_rows = len(ROWS)
# header (80) + date (50) + warn (50) + col header (36) + rows + group headers + footer (50) + padding
n_groups = len(set(r[2] for r in ROWS))
H = 80 + 50 + (50 if ANN_OFF else 0) + 36 + total_rows * ROW_H + n_groups * ROW_H + 50 + PAD

img = Image.new("RGB", (W, H), BG)
d   = ImageDraw.Draw(img)

def pill(draw, x, y, text, bg, fg, font):
    tw, th = draw.textbbox((0,0), text, font=font)[2:]
    rx, ry, rw, rh = x, y - 1, tw + 14, th + 6
    draw.rounded_rectangle([rx, ry, rx+rw, ry+rh], radius=4, fill=bg)
    draw.text((rx + 7, ry + 3), text, font=font, fill=fg)
    return rw

def hline(draw, y, color=BORDER, x0=PAD, x1=None):
    draw.line([(x0, y), (x1 or W - PAD, y)], fill=color, width=1)

y = 0

# ── HEADER BAR ───────────────────────────────────────────────────────────────
d.rectangle([0, 0, W, 80], fill=HEADER_BG)
d.text((PAD, 14), "🍽  LUNCH SCHEDULE", font=F_TIT, fill=HEADER_FG)
d.text((PAD, 48), DATE, font=F_SM, fill="#BFB9AC")
y = 80

# ── DATE / SUBTITLE ──────────────────────────────────────────────────────────
d.rectangle([0, y, W, y + 50], fill="#FEFCF7")
d.text((PAD, y + 14), "EST window: 11:00 AM – 4:00 PM  ·  PST fixed: 4:00 PM PST  ·  Sorted by lead volume (high → early)", font=F_XS, fill=MU)
hline(d, y + 50)
y += 50

# ── WARNING (if flex VA is off) ───────────────────────────────────────────────
if ANN_OFF:
    d.rectangle([0, y, W, y + 50], fill=WARN_BG)
    d.text((PAD, y + 14), "⚠  Ann is OFF today — Ed, Mark, Myra redistributed to Maykee & Reyn based on lead volume balance.", font=F_SMB, fill=WARN_FG)
    hline(d, y + 50, WARN_FG)
    y += 50

# ── COLUMN HEADERS ───────────────────────────────────────────────────────────
d.rectangle([0, y, W, y + 36], fill=SUBHD_BG)
cx = PAD
for i, (col, cw) in enumerate(zip(COLS, COL_W)):
    d.text((cx + 4, y + 8), col, font=F_XS, fill=MU)
    cx += cw
hline(d, y + 36, "#B0A99C")
y += 36

# ── ROWS ─────────────────────────────────────────────────────────────────────
prev_group = None
for idx, (name, lunch, covered, tz, moved) in enumerate(ROWS):
    # Group divider
    if covered != prev_group:
        if prev_group is not None:
            hline(d, y, "#B0A99C")
        g_bg  = MAYKEE_BG if covered == "Maykee" else REYN_BG
        g_fg  = MAYKEE_FG if covered == "Maykee" else REYN_FG
        label = f"  {covered.upper()} GROUP"
        d.rectangle([0, y, W, y + ROW_H], fill=g_bg)
        d.text((PAD, y + 11), label, font=F_SMB, fill=g_fg)
        hline(d, y + ROW_H, g_fg + "66" if len(g_fg) == 7 else g_fg)
        y += ROW_H
        prev_group = covered

    bg = ROW_MOVED if moved else (ROW_ALT if idx % 2 == 1 else ROW_NORM)
    d.rectangle([0, y, W, y + ROW_H], fill=bg)

    cx = PAD
    # VA name
    name_font = F_BOL if moved else F_REG
    d.text((cx + 4, y + 10), name, font=name_font, fill=DARK)
    if moved:
        pill(d, cx + 4 + d.textbbox((0,0), name, font=F_BOL)[2] + 8, y + 12, "moved", MOVED_BG, MOVED_FG, F_XS)
    cx += COL_W[0]

    # Lunch time
    is_pst = tz == "PST"
    d.text((cx + 4, y + 10), lunch + (" PST" if is_pst else ""), font=F_REG, fill=DARK)
    cx += COL_W[1]

    # Covered by
    d.text((cx + 4, y + 10), covered, font=F_REG, fill=MAYKEE_FG if covered == "Maykee" else REYN_FG)
    cx += COL_W[2]

    # TZ badge
    pill(d, cx + 4, y + 11, tz, PST_BG if is_pst else EST_BG, PST_FG if is_pst else EST_FG, F_XS)

    hline(d, y + ROW_H)
    y += ROW_H

# ── FOOTER ───────────────────────────────────────────────────────────────────
d.rectangle([0, y, W, H], fill=HEADER_BG)
d.text((PAD, y + 12), "Roof Ignite  ·  VA Calendar", font=F_XS, fill="#BFB9AC")

out = "/home/user/va-calendar/lunch_schedule_2026-05-22.png"
img.save(out, "PNG")
print(f"Saved: {out}  ({W}×{H}px)")
