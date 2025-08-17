# app.py
# pip install streamlit reportlab pandas

import os, io, math
import pandas as pd
import streamlit as st

from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.colors import Color, black
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ======== Fixed look & feel ========
FONT_REQ        = "ArialBlack"
FONT_SIZE_PT    = 18
FONT_MIN_PT     = 12
TEXT_COLOR      = Color(128/255, 0, 0)   # maroon
DRAW_BORDERS    = True
BORDER_COLOR    = black
LINE_COLOR      = black
LINE_THICK_PT   = 1.0

# Divider & text padding (inches)
INNER_PAD_IN  = 0.05
V_TEXT_PAD_IN = 0.06
H_TEXT_PAD_IN = 0.06

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

def draw_centered_text_in_region(c, x_center_pt, y0_pt, y1_pt, text, font_name, base_font_size_pt, max_width_pt, color):
    font_size_pt = fit_font_size(font_name, text, max_width_pt, base_font_size_pt, FONT_MIN_PT)
    c.setFillColor(color)
    c.setFont(font_name, font_size_pt)
    w = c.stringWidth(text, font_name, font_size_pt)
    a = pdfmetrics.getAscent(font_name)
    d = pdfmetrics.getDescent(font_name)  # negative
    region_center = (y0_pt + y1_pt) / 2.0
    baseline = region_center - ((a + d) * font_size_pt) / 2000.0
    c.drawString(x_center_pt - w/2.0, baseline, text)

def draw_sticker(c, x_pt, y_pt, w_pt, h_pt, top_text, bottom_text,
                 font_name, font_size_pt, text_color, draw_border=True):
    if draw_border:
        c.setStrokeColor(BORDER_COLOR)
        c.rect(x_pt, y_pt, w_pt, h_pt, stroke=1, fill=0)

    # Divider at halfway
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
        draw_centered_text_in_region(c, x_center, top_y0, top_y1, str(top_text),
                                     font_name, font_size_pt, max_text_w, text_color)
    if bottom_text:
        draw_centered_text_in_region(c, x_center, bot_y0, bot_y1, str(bottom_text),
                                     font_name, font_size_pt, max_text_w, text_color)

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
    """
    df = jobs_df.copy().fillna("")
    df["top"] = df["top"].astype(str).str.strip()
    df["bottom"] = df["bottom"].astype(str).str.strip()
    df["count"] = pd.to_numeric(df["count"], errors="coerce").fillna(0).astype(int)
    df = df[(df["top"] != "") & (df["bottom"] != "") & (df["count"] > 0)]
    if df.empty:
        raise ValueError("No valid sticker rows. Fill top, bottom, and a positive count.")

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
            for _ in range(int(row["count"])):
                yield (row["top"], row["bottom"])

    stream = sticker_stream()
    total_count = int(df["count"].sum())
    drawn = 0

    while True:
        for slot in range(capacity):
            try:
                top, bottom = next(stream)
                is_blank = False
            except StopIteration:
                top, bottom = "", ""
                is_blank = True

            # Row-major, TOP-LEFT first:
            r = slot // cols                 # 0..rows-1 (top-first logical row)
            col = slot % cols                # 0..cols-1 (left→right)
            row_from_bottom = (rows - 1) - r
            x = left_margin_pt + offset_x_pt + col * sticker_w_pt
            y = bottom_margin_pt + offset_y_pt + row_from_bottom * sticker_h_pt

            draw_sticker(c, x, y, sticker_w_pt, sticker_h_pt,
                         top, bottom, font_name, FONT_SIZE_PT, TEXT_COLOR, DRAW_BORDERS)
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

# Live capacity preview
work_w = page_w_in - 2*margin_in
work_h = page_h_in - 2*margin_in
cols, rows, capacity = compute_grid(work_w, work_h, sticker_w_in, sticker_h_in)
st.caption(f"Working area: **{work_w:.3f} × {work_h:.3f} in** · Grid: **{cols} × {rows} = {capacity}** per page")

# Session state: list of stickers [{top, bottom, count}, ...]
if "stickers" not in st.session_state:
    st.session_state.stickers = [{"top": "5001", "bottom": "19608", "count": 1}]
if "show_add_form" not in st.session_state:
    st.session_state.show_add_form = False

# Add sticker flow (mobile-friendly)
if not st.session_state.show_add_form:
    if st.button("➕ Add sticker", use_container_width=True):
        st.session_state.show_add_form = True
else:
    with st.form("add_sticker_form", clear_on_submit=True):
        top_val = st.text_input("Top (upper value)")
        bottom_val = st.text_input("Bottom (lower value)")
        count_val = st.number_input("Number of stickers", min_value=1, step=1, value=1)
        c1, c2 = st.columns(2)
        add_clicked = c1.form_submit_button("Add", use_container_width=True)
        cancel_clicked = c2.form_submit_button("Cancel", use_container_width=True)

    if add_clicked:
        t = (top_val or "").strip()
        b = (bottom_val or "").strip()
        if not t or not b:
            st.warning("Please enter both Top and Bottom values.")
        else:
            st.session_state.stickers.append({"top": t, "bottom": b, "count": int(count_val)})
            st.session_state.show_add_form = False
    elif cancel_clicked:
        st.session_state.show_add_form = False

# Current list view (simple + removable)
st.subheader("Stickers to print (in order)")
if not st.session_state.stickers:
    st.info("No stickers added yet.")
else:
    to_remove = None
    for i, item in enumerate(st.session_state.stickers):
        with st.container(border=True):
            st.write(f"**{i+1}.** Top: `{item['top']}` · Bottom: `{item['bottom']}` · Count: `{item['count']}`")
            if st.button("Remove", key=f"rm_{i}"):
                to_remove = i
    if to_remove is not None:
        st.session_state.stickers.pop(to_remove)

col_left, col_right = st.columns(2)
with col_left:
    if st.button("🧹 Clear all", use_container_width=True):
        st.session_state.stickers = []

st.divider()
generate = st.button("📄 Generate PDF", type="primary", use_container_width=True)

if generate:
    try:
        # Convert list → DataFrame for the PDF function
        jobs_df = pd.DataFrame(st.session_state.stickers)
        pdf_bytes, summary = make_multi_sticker_pdf_dynamic(
            jobs_df, page_w_in, page_h_in, margin_in, sticker_w_in, sticker_h_in
        )
        st.success(
            f"Generated PDF with {summary['total']} stickers across {summary['pages']} page(s). "
            f"(Per page: {summary['cols']} × {summary['rows']} = {summary['per_page']})"
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
