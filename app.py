# app.py
# pip install streamlit reportlab pandas openpyxl

import os, io, math
import pandas as pd
import streamlit as st

from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.colors import Color, black, green, yellow, red, blue, purple, orange, brown
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ======== Fixed look & feel ========
FONT_REQ        = "ArialBlack"
FONT_SIZE_PT    = 18
FONT_MIN_PT     = 12
# Pre-defined color palette for cycling through different colors
COLOR_PALETTE = [
    Color(128/255, 0, 0),        # maroon
    black,                       # black
    green,                       # green
    yellow,                      # yellow
    red,                         # red
    blue,                        # blue
    purple,                      # purple
    orange,                      # orange
    brown,                       # brown
    Color(0, 128/255, 128/255), # teal
    Color(128/255, 0, 128/255), # purple
    Color(0, 0, 128/255),       # navy
]
DRAW_BORDERS    = False                  # keep only the middle divider
BORDER_COLOR    = black
LINE_COLOR      = black
LINE_THICK_PT   = 1.0

# Divider & text padding (inches)
INNER_PAD_IN  = 0.05
V_TEXT_PAD_IN = 0.06
H_TEXT_PAD_IN = 0.06

# Fine-tune vertical centering (points). Negative = move down; positive = move up
# Tip: set to -6 to move the TOP value slightly closer to the center line.
TOP_CENTER_BIAS_PT = -2

# ======== Font helpers ========
def try_register_arial_black() -> str:
    """Try to register Arial Black from repo or OS; fallback to Helvetica-Bold."""
    candidates = [
        os.path.join(os.getcwd(), "ariblk.ttf"),
        os.path.join(os.getcwd(), "ARIBLK.TTF"),
        os.path.join(os.getcwd(), "fonts", "ariblk.ttf"),
        os.path.join(os.getcwd(), "fonts", "ARIBLK.TTF"),
        r"C:\Windows\Fonts\ariblk.ttf",
        r"C:\Windows\Fonts\ARIBLK.TTF",
        r"/Library/Fonts/Arial Black.ttf",
        r"/System/Library/Fonts/Supplemental/Arial Black.ttf",
    ]
    for path in candidates:
        if os.path.isfile(path):
            try:
                pdfmetrics.registerFont(TTFont(FONT_REQ, path))
                return FONT_REQ
            except Exception:
                pass
    return "Helvetica-Bold"  # fallback

def fit_font_size(font_name, text, max_width_pt, desired_pt, min_pt):
    w = pdfmetrics.stringWidth(text, font_name, desired_pt)
    if w <= max_width_pt:
        return desired_pt
    if desired_pt <= min_pt:
        return min_pt
    return max(min_pt, desired_pt * (max_width_pt / w))

def draw_centered_text_in_region(
    c, x_center_pt, y0_pt, y1_pt, text, font_name,
    base_font_size_pt, max_width_pt, color, center_bias_pt=0.0
):
    font_size_pt = fit_font_size(font_name, text, max_width_pt, base_font_size_pt, FONT_MIN_PT)
    c.setFillColor(color)
    c.setFont(font_name, font_size_pt)
    w = c.stringWidth(text, font_name, font_size_pt)
    a = pdfmetrics.getAscent(font_name)
    d = pdfmetrics.getDescent(font_name)  # negative
    region_center = (y0_pt + y1_pt) / 2.0
    baseline = region_center - ((a + d) * font_size_pt) / 2000.0 + center_bias_pt
    c.drawString(x_center_pt - w/2.0, baseline, text)

def draw_sticker(c, x_pt, y_pt, w_pt, h_pt, top_text, bottom_text,
                 font_name, font_size_pt, text_color, draw_border=False):
    # Outer border removed by default
    if draw_border:
        c.setStrokeColor(BORDER_COLOR)
        c.rect(x_pt, y_pt, w_pt, h_pt, stroke=1, fill=0)

    # Divider at halfway (we keep only this middle line)
    line_y = y_pt + h_pt * 0.5
    line_pad = INNER_PAD_IN * inch
    c.setStrokeColor(LINE_COLOR)
    c.setLineWidth(LINE_THICK_PT)
    c.line(x_pt + line_pad, line_y, x_pt + w_pt - line_pad, line_y)

    # Text regions + padding
    vpad = V_TEXT_PAD_IN * inch
    hpad = H_TEXT_PAD_IN * inch
    top_y0, top_y1 = line_y + vpad, (y_pt + h_pt) - vpad
    bot_y0, bot_y1 = y_pt + vpad, line_y - vpad
    x_center = x_pt + w_pt / 2.0
    max_text_w = w_pt - 2*hpad

    if top_text:
        draw_centered_text_in_region(
            c, x_center, top_y0, top_y1, str(top_text),
            font_name, font_size_pt, max_text_w, text_color,
            center_bias_pt=TOP_CENTER_BIAS_PT
        )
    if bottom_text:
        draw_centered_text_in_region(
            c, x_center, bot_y0, bot_y1, str(bottom_text),
            font_name, font_size_pt, max_text_w, text_color,
            center_bias_pt=0.0
        )

def compute_grid(work_w_in, work_h_in, sticker_w_in, sticker_h_in):
    """Max whole stickers that fit (no gaps)."""
    cols = max(int(math.floor(work_w_in / sticker_w_in)), 0)
    rows = max(int(math.floor(work_h_in / sticker_h_in)), 0)
    return cols, rows, cols * rows

# ======== PDF writer (dynamic grid per specs, pad last page) ========
def make_multi_sticker_pdf_dynamic(
    jobs_df: pd.DataFrame,
    page_w_in: float, page_h_in: float, margin_in: float,
    sticker_w_in: float, sticker_h_in: float,
):
    """
    - Grid = floor(working_area / sticker_size)
    - Constant per-page capacity across all pages
    - Pad last page from top-left so it still has full grid with trailing blanks
    - Each unique row gets a different color from the palette
    """
    df = jobs_df.copy().fillna("")
    # Convert to string to accept both numbers and text
    df["top"] = df["top"].astype(str).str.strip()
    df["bottom"] = df["bottom"].astype(str).str.strip()
    df["count"] = pd.to_numeric(df["count"], errors="coerce").fillna(0).astype(int)
    df = df[(df["top"] != "") & (df["bottom"] != "") & (df["count"] > 0)]
    if df.empty:
        raise ValueError("No valid sticker rows. Fill top, bottom, and a positive count.")

    # Assign a color to each unique row (top+bottom combination)
    unique_rows = df[['top', 'bottom']].drop_duplicates()
    color_mapping = {}
    for i, (_, row) in enumerate(unique_rows.iterrows()):
        color_index = i % len(COLOR_PALETTE)
        color_mapping[(row['top'], row['bottom'])] = COLOR_PALETTE[color_index]

    work_w_in = page_w_in - 2 * margin_in
    work_h_in = page_h_in - 2 * margin_in
    if work_w_in <= 0 or work_h_in <= 0:
        raise ValueError("Margins are too large for the page size.")

    cols, rows, capacity = compute_grid(work_w_in, work_h_in, sticker_w_in, sticker_h_in)
    if capacity <= 0:
        raise ValueError("Sticker size does not fit the working area. Increase page, reduce margins, or shrink stickers.")

    # points
    page_w_pt = page_w_in * inch
    page_h_pt = page_h_in * inch
    left_margin_pt   = margin_in * inch
    bottom_margin_pt = margin_in * inch
    sticker_w_pt = sticker_w_in * inch
    sticker_h_pt = sticker_h_in * inch

    # center the grid inside working area
    used_w_pt = cols * sticker_w_pt
    used_h_pt = rows * sticker_h_pt
    work_w_pt = work_w_in * inch
    work_h_pt = work_h_in * inch
    offset_x_pt = (work_w_pt - used_w_pt) / 2.0
    offset_y_pt = (work_h_pt - used_h_pt) / 2.0

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_w_pt, page_h_pt))
    font_name = try_register_arial_black()

    # order preserved exactly as entered
    def sticker_stream():
        for _, row in df.iterrows():
            color = color_mapping[(row["top"], row["bottom"])]
            for _ in range(int(row["count"])):
                yield (row["top"], row["bottom"], color)

    stream = sticker_stream()
    total_count = int(df["count"].sum())
    drawn = 0

    while True:
        for slot in range(capacity):
            try:
                top, bottom, color = next(stream)
                is_blank = False
            except StopIteration:
                top, bottom, color = "", "", black
                is_blank = True

            # Row-major, TOP-LEFT first:
            r = slot // cols
            col = slot % cols
            row_from_bottom = (rows - 1) - r
            x = left_margin_pt + offset_x_pt + col * sticker_w_pt
            y = bottom_margin_pt + offset_y_pt + row_from_bottom * sticker_h_pt

            draw_sticker(
                c, x, y, sticker_w_pt, sticker_h_pt,
                top, bottom, font_name, FONT_SIZE_PT, color, DRAW_BORDERS
            )
            if not is_blank:
                drawn += 1

        c.showPage()
        if drawn >= total_count:
            break

    c.save()
    pdf_bytes = buf.getvalue()
    buf.close()

    return pdf_bytes, {
        "cols": cols, "rows": rows, "per_page": capacity,
        "total": drawn, "pages": math.ceil(drawn / capacity),
        "colors_used": len(set(color_mapping.values()))
    }

# ========================= Streamlit UI =========================
st.set_page_config(page_title="Sticker PDF Generator", page_icon="🖨️", layout="centered")
st.title("🖨️ Sticker PDF Generator")

with st.sidebar:
    st.header("Specs (inches)")
    page_w_in    = st.number_input("Page width",  min_value=1.0, value=12.0, step=0.1, format="%.3f")
    page_h_in    = st.number_input("Page height", min_value=1.0, value=18.0, step=0.1, format="%.3f")
    margin_in    = st.number_input("Margin (all sides)", min_value=0.0, value=0.25, step=0.05, format="%.3f")
    sticker_w_in = st.number_input("Sticker width",  min_value=0.1, value=1.134, step=0.01, format="%.3f")
    sticker_h_in = st.number_input("Sticker height", min_value=0.1, value=0.585, step=0.01, format="%.3f")

# Upload input file (Excel/CSV) with exactly 3 columns: Top, Bottom, Count
st.subheader("Input file")
uploaded = st.file_uploader("Upload Excel/CSV (3 columns: Top, Bottom, Count)", type=["xlsx", "xls", "csv"])

# Keep a flag to avoid multi-click issues
if "do_generate" not in st.session_state:
    st.session_state.do_generate = False

def _load_df_from_upload(file) -> pd.DataFrame:
    # Accept both Excel and CSV; assume the first 3 columns are Top/Bottom/Count
    name = file.name.lower()
    if name.endswith(".csv"):
        df_raw = pd.read_csv(file, header=None)  # simplest: 3 cols in order
    else:
        df_raw = pd.read_excel(file, header=None, engine="openpyxl")

    if df_raw.shape[1] < 3:
        raise ValueError("The uploaded file must have at least 3 columns (Top, Bottom, Count).")

    df = pd.DataFrame({
        "top":    df_raw.iloc[:, 0],
        "bottom": df_raw.iloc[:, 1],
        "count":  df_raw.iloc[:, 2],
    })
    # Trim/clean - convert to string to accept both numbers and text
    df["top"] = df["top"].astype(str).str.strip()
    df["bottom"] = df["bottom"].astype(str).str.strip()
    df["count"] = pd.to_numeric(df["count"], errors="coerce").fillna(0).astype(int)
    # Keep only valid rows
    df = df[(df["top"] != "") & (df["bottom"] != "") & (df["count"] > 0)].reset_index(drop=True)
    if df.empty:
        raise ValueError("No valid rows found. Ensure the first 3 columns are Top, Bottom, Count with positive counts.")
    return df

# Preview + Generate
if uploaded is not None:
    try:
        jobs_df = _load_df_from_upload(uploaded)
        st.success(f"Loaded {len(jobs_df)} sticker rows from file.")
        st.dataframe(jobs_df.head(50), use_container_width=True)

        # Generate button sets the flag; generation runs in the next block
        def _on_generate():
            st.session_state.do_generate = True

        st.button("📄 Generate PDF", type="primary", use_container_width=True, on_click=_on_generate)

    except Exception as e:
        st.error(str(e))

# If flagged, generate once on this rerun
if st.session_state.do_generate and uploaded is not None:
    try:
        jobs_df = _load_df_from_upload(uploaded)  # reload to be safe
        pdf_bytes, summary = make_multi_sticker_pdf_dynamic(
            jobs_df, page_w_in, page_h_in, margin_in, sticker_w_in, sticker_h_in
        )
        st.success(
            f"Generated PDF with {summary['total']} stickers across {summary['pages']} page(s). "
            f"(Per page: {summary['cols']} × {summary['rows']} = {summary['per_page']}) "
            f"Using {summary['colors_used']} different colors."
        )
        st.download_button(
            label="⬇️ Download sticker_sheet.pdf",
            data=pdf_bytes,
            file_name="sticker_sheet.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    except Exception as e:
        st.error(str(e))
    finally:
        st.session_state.do_generate = False

# Helper caption
st.caption("File format: exactly three columns per row → [Top, Bottom, Count]. No headers needed. "
           "Top and Bottom can contain both numbers and text. Each unique row gets a different color.")
