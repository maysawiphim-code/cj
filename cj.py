import streamlit as st
import pandas as pd
import json, re, io, urllib.request
import numpy as np

# matplotlib ต้อง set backend ก่อน import อื่นๆ
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="CJ Smart Scan", page_icon="🛒", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;500;600;700;800&display=swap');
@import url('https://fonts.googleapis.com/css2?family=Fredoka+One&display=swap');
*, html, body, [class*="css"] { font-family: 'Sarabun', sans-serif !important; }
#MainMenu, header, footer { visibility: hidden; }
.stApp { background: linear-gradient(160deg,#FFF8F0 0%,#FFF0F5 50%,#F0F4FF 100%); }
.hero {
    background: linear-gradient(135deg,#FF6B6B 0%,#FF4D6D 40%,#C9184A 100%);
    border-radius: 24px; padding: 36px 40px 32px; margin-bottom: 32px;
    position: relative; overflow: hidden;
    box-shadow: 0 8px 32px rgba(201,24,74,.25);
}
.hero-title {
    font-family:'Fredoka One','Sarabun',sans-serif !important;
    font-size:2.6rem; font-weight:900; color:white; margin:0;
    text-shadow:0 2px 8px rgba(0,0,0,.2);
}
.hero-sub { color:rgba(255,255,255,.88); font-size:1.05rem; margin:10px 0 0; }
.hero-mascot {
    position:absolute; right:40px; top:50%; transform:translateY(-50%);
    font-size:7rem; filter:drop-shadow(0 4px 12px rgba(0,0,0,.15));
}
.sec-title {
    display:flex; align-items:center; gap:10px;
    font-size:1.15rem; font-weight:700; color:#C9184A; margin:28px 0 14px;
}
.sec-title .dot { width:6px; height:24px; background:#FF4D6D; border-radius:3px; }
.mcard {
    background:white; border-radius:18px; padding:22px 18px; text-align:center;
    box-shadow:0 4px 16px rgba(0,0,0,.07); transition:transform .15s;
}
.mcard:hover { transform:translateY(-3px); }
.mcard .icon { font-size:2.4rem; line-height:1; margin-bottom:6px; }
.mcard .num  { font-size:2rem; font-weight:800; color:#C9184A; }
.mcard .lbl  { font-size:.82rem; color:#888; margin-top:4px; font-weight:500; }
.cat-card {
    background:white; border-radius:16px; padding:16px 20px; margin-bottom:10px;
    display:flex; align-items:center; gap:14px;
    box-shadow:0 2px 10px rgba(0,0,0,.06);
}
.cat-card .cat-icon { font-size:2rem; }
.cat-card .cat-name  { font-weight:700; font-size:.95rem; color:#333; }
.cat-card .cat-count { font-size:.82rem; color:#888; }
.cat-card .bar-wrap  { flex:1; background:#f5f5f5; border-radius:99px; height:8px; overflow:hidden; }
.cat-card .bar-fill  { height:100%; border-radius:99px; }
.cat-card .cat-num   { font-weight:800; font-size:1.1rem; color:#C9184A; min-width:36px; text-align:right; }
.stDownloadButton > button {
    background:linear-gradient(135deg,#FF6B6B,#C9184A) !important;
    color:white !important; border:none !important; border-radius:14px !important;
    padding:14px 28px !important; font-size:1rem !important; font-weight:700 !important;
    box-shadow:0 4px 16px rgba(201,24,74,.3) !important;
}
.stButton > button[kind="primary"] {
    background:linear-gradient(135deg,#FF6B6B,#C9184A) !important;
    color:white !important; border:none !important; border-radius:14px !important;
    padding:14px 28px !important; font-size:1rem !important; font-weight:700 !important;
    box-shadow:0 4px 16px rgba(201,24,74,.3) !important;
}
.footer { text-align:center; color:#bbb; font-size:.8rem; padding:32px 0 16px; }
</style>
""", unsafe_allow_html=True)

# ─── CONSTANTS ───
CATEGORIES = {
    "Bao Cafe":         {"icon":"☕","color":"#C2185B","en":"Bao Cafe"},
    "อาหารพร้อมทาน":   {"icon":"🍱","color":"#F57C00","en":"Ready-to-eat"},
    "เครื่องดื่ม":      {"icon":"🥤","color":"#1976D2","en":"Beverages"},
    "ขนมขบเคี้ยว":     {"icon":"🍿","color":"#7B1FA2","en":"Snacks"},
    "ของใช้ส่วนตัว":   {"icon":"🧴","color":"#00897B","en":"Personal Care"},
    "ของใช้ในบ้าน":    {"icon":"🏠","color":"#558B2F","en":"Household"},
    "สินค้าเบ็ดเตล็ด": {"icon":"🔋","color":"#6D4C41","en":"Misc"},
    "บริการและอื่นๆ":  {"icon":"📱","color":"#455A64","en":"Services"},
}

SYSTEM_PROMPT = """คุณคือระบบจำแนกประเภทสินค้าร้านสะดวกซื้อ CJ Express
วิเคราะห์จากชื่อสินค้าเท่านั้น โดยพิจารณาจากความหมาย ลักษณะ และประเภทของสินค้าที่ชื่อบ่งบอก

หมวดหมู่ที่ใช้ได้:
1. Bao Cafe        → เครื่องดื่มจากร้าน Bao Cafe เช่น กาแฟ ลาเต้ เอสเปรสโซ่ อเมริกาโน่ มัทฉะ ชาเขียวนมสด (ชื่อมักขึ้นต้นด้วย Bao_ หรือ Bac_)
2. อาหารพร้อมทาน  → อาหารพร้อมรับประทาน เช่น แซนด์วิช ปัง เบเกอรี่ ไส้กรอก อาหารแช่แข็ง พุดดิ้ง ข้าวเกรียบ ผงชูรส ไข่ มินิบัน แฮม
3. เครื่องดื่ม     → เครื่องดื่มทั่วไป เช่น น้ำดื่ม น้ำอัดลม นม ชา กาแฟสำเร็จรูป เบียร์ สุรา น้ำหวาน โซดา เครื่องดื่มชูกำลัง คอลลาเจน โพรไบโอติกส์
4. ขนมขบเคี้ยว    → ขนมและของกินเล่น เช่น มันฝรั่ง ลูกอม ช็อกโกแลต ไอศกรีม ถั่ว เมล็ดทานตะวัน อัลมอนด์ แครกเกอร์ ข้าวโพดอบ ขนมกรอบ
5. ของใช้ส่วนตัว  → ผลิตภัณฑ์ดูแลร่างกาย เช่น ยาสีฟัน แชมพู สบู่ ครีม ผ้าอนามัย โฟมล้างหน้า เครื่องสำอาง รองพื้น แปรงสีฟัน ยา วิตามิน เจล
6. ของใช้ในบ้าน   → ผลิตภัณฑ์ทำความสะอาด เช่น ทิชชู ผงซักฟอก น้ำยาล้างจาน ถุงขยะ น้ำยาทำความสะอาด น้ำยาปรับผ้านุ่ม
7. สินค้าเบ็ดเตล็ด → สินค้าทั่วไปอื่นๆ เช่น ถ่านไฟฉาย ไฟแช็ก ยากันยุง เครื่องเขียน เทปลบคำผิด ยางลบ
8. บริการและอื่นๆ  → บริการและส่วนลด เช่น เติมเงินมือถือ จ่ายบิล ซิมการ์ด บัตรเติมเงิน ส่วนลด แก้วลูกค้า

กฎสำคัญ:
- วิเคราะห์จากชื่อสินค้าล้วนๆ ไม่ใช้ข้อมูลอื่น
- ตอบเป็น JSON บรรทัดเดียว ไม่มี markdown ไม่มีข้อความนำหน้า
- รูปแบบ: {"ชื่อสินค้า1":"หมวดหมู่1","ชื่อสินค้า2":"หมวดหมู่2"}
- ถ้าชื่อสินค้าไม่ชัดเจนให้วิเคราะห์จากส่วนที่อ่านได้ ถ้าไม่แน่ใจจริงๆ ใส่ "สินค้าเบ็ดเตล็ด"
"""

# ─── CHART GENERATOR ───
RANK_COLORS = ["#FFD700","#C0C0C0","#CD7F32","#FF6B6B","#FF9F43",
               "#48DBFB","#1DD1A1","#A29BFE","#FD79A8","#636E72"]

def draw_cell(ax, x, y, w, h, text, bg, fg, bold=False, fs=10):
    rect = FancyBboxPatch((x+0.002, y+0.005), w-0.004, h-0.008,
                           boxstyle="round,pad=0.008", linewidth=0,
                           facecolor=bg, clip_on=False)
    ax.add_patch(rect)
    ax.text(x+w/2, y+h/2, str(text), ha="center", va="center",
            fontsize=fs, fontweight="bold" if bold else "normal",
            color=fg, clip_on=False)

def make_top_branch_table(items_df):
    """สร้าง figure ตาราง top10 สาขาขายดีรายเดือน"""
    df = items_df.copy()
    df["วันที่_dt"] = pd.to_datetime(df["วันที่"], format="%d/%m/%Y", errors="coerce")
    df["เดือน"]     = df["วันที่_dt"].dt.strftime("%b %Y")
    df["เดือน_sort"]= df["วันที่_dt"].dt.to_period("M").astype(str)

    # รวม top category ต่อสาขาต่อเดือน
    cat_grp = df.groupby(["เดือน_sort","เดือน","รหัสสาขา","ประเภทสินค้า"])["ชื่อสินค้า"].count()
    top_cat = cat_grp.reset_index().sort_values("ชื่อสินค้า", ascending=False)
    top_cat = top_cat.drop_duplicates(subset=["เดือน_sort","รหัสสาขา"]).rename(
        columns={"ชื่อสินค้า":"_cnt","ประเภทสินค้า":"top_cat"})

    grp = df.groupby(["เดือน_sort","เดือน","รหัสสาขา"])
    cnt = grp["ชื่อสินค้า"].count()
    tot = grp["ยอดรวมสินค้า"].sum()
    monthly = pd.concat([cnt, tot], axis=1)
    monthly.columns = ["จำนวนรายการ","ยอดรวม"]
    monthly = monthly.reset_index()
    monthly = monthly.merge(
        top_cat[["เดือน_sort","รหัสสาขา","top_cat"]],
        on=["เดือน_sort","รหัสสาขา"], how="left")

    months = sorted(monthly["เดือน_sort"].unique())
    n_months = len(months)

    fig_h = max(5, n_months * 3.8)
    fig = plt.figure(figsize=(14, fig_h), facecolor="#FFF8F0")
    fig.suptitle("Top 10 Branches by Monthly Sales",
                 fontsize=16, fontweight="bold", color="#C9184A", y=0.99)

    col_labels  = ["Rank","Branch ID","Month","Total Sales","Items","Top Category"]
    col_widths  = [0.065, 0.13, 0.14, 0.20, 0.095, 0.24]
    cw_sum      = sum(col_widths)
    cx          = [sum(col_widths[:i])/cw_sum for i in range(len(col_widths))]
    cw          = [w/cw_sum for w in col_widths]
    H = 0.075   # row height in axes coords

    for m_idx, month_key in enumerate(months):
        mdata = monthly[monthly["เดือน_sort"]==month_key]                .sort_values("ยอดรวม", ascending=False).head(10)
        month_label = mdata["เดือน"].iloc[0]
        n_rows = len(mdata)
        total_h = (n_rows+1)*H + 0.08   # header + rows + title

        top = 1 - m_idx*(total_h + 0.06) - 0.04
        bot = top - total_h
        if bot < 0.01: break

        ax = fig.add_axes([0.02, bot, 0.96, total_h])
        ax.set_xlim(0,1); ax.set_ylim(0,1); ax.axis("off")
        ax.set_facecolor("#FFF8F0")

        # ── Month title badge ──
        ax.text(0.0, 1.0, f" 📅 {month_label}",
                ha="left", va="bottom", fontsize=12, fontweight="bold",
                color="#C9184A", transform=ax.transAxes)

        # ── Header ──
        hy = 1 - 0.06/total_h - H
        for j in range(len(col_labels)):
            draw_cell(ax, cx[j], hy, cw[j], H,
                      col_labels[j], "#C9184A", "white", bold=True, fs=9.5)

        # ── Data rows ──
        for i, (_, row) in enumerate(mdata.iterrows()):
            y = hy - (i+1)*H
            bg = "#FFF0F3" if i%2==0 else "#FFFFFF"
            rc  = RANK_COLORS[i] if i < len(RANK_COLORS) else "#AAAAAA"
            cat = str(row.get("top_cat",""))
            cat_color = CATEGORIES.get(cat,{"color":"#9E9E9E"})["color"]
            cat_en    = CATEGORIES.get(cat,{"en":cat})["en"]
            vals = [
                str(i+1),
                str(int(row["รหัสสาขา"])),
                str(row["เดือน"]),
                f'{row["ยอดรวม"]:,.0f}',
                str(int(row["จำนวนรายการ"])),
                cat_en,
            ]
            for j, val in enumerate(vals):
                if j == 0:
                    draw_cell(ax, cx[j], y, cw[j], H, val, rc, "white", bold=True, fs=10)
                elif j == len(vals)-1:
                    draw_cell(ax, cx[j], y, cw[j], H, val, cat_color+"33", cat_color, bold=True, fs=9)
                else:
                    draw_cell(ax, cx[j], y, cw[j], H, val, bg, "#333", fs=9.5)

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor="#FFF8F0")
    plt.close(fig)
    buf.seek(0)
    return buf

def make_chart(summary_df, items_df=None):
    cats   = summary_df["ประเภทสินค้า"].tolist()
    counts = summary_df["จำนวนรายการ"].tolist()
    colors = [CATEGORIES.get(c, {"color":"#9E9E9E"})["color"] for c in cats]
    labels = [CATEGORIES.get(c, {"en":c})["en"] for c in cats]

    fig = plt.figure(figsize=(18, 9), facecolor="#FFF8F0")
    fig.suptitle("CJ Express — Product Category Analysis",
                 fontsize=20, fontweight="bold", color="#C9184A", y=0.98)

    # ── Bar chart ──
    ax1 = fig.add_axes([0.04, 0.10, 0.44, 0.82])
    ax1.set_facecolor("#FFF8F0")
    idx = np.argsort(counts)
    s_lbl = [labels[i] for i in idx]
    s_cnt = [counts[i] for i in idx]
    s_col = [colors[i] for i in idx]
    y_pos = np.arange(len(s_lbl))

    for i, (lbl, cnt, col) in enumerate(zip(s_lbl, s_cnt, s_col)):
        ax1.barh(i, cnt, height=0.52, color=col, linewidth=0, zorder=3)
        ax1.barh(i+0.1, cnt*0.9, height=0.14, color="white", alpha=0.22, left=0.2, zorder=4)
        ax1.text(cnt+0.3, i, str(cnt), va="center", fontsize=13,
                 fontweight="bold", color=col, zorder=5)
        ax1.scatter(-0.8, i, s=380, c=col, zorder=6, linewidths=2, edgecolors="white")

    ax1.set_yticks(y_pos)
    ax1.set_yticklabels(s_lbl, fontsize=10.5)
    ax1.set_xlim(-1.6, max(counts)*1.22)
    ax1.set_xlabel("Number of Items", fontsize=10, color="#777")
    ax1.set_title("Items per Category", fontsize=13, fontweight="bold", color="#333", pad=10)
    ax1.spines["top"].set_visible(False)
    ax1.spines["right"].set_visible(False)
    ax1.spines["left"].set_visible(False)
    ax1.tick_params(left=False, colors="#555")
    ax1.grid(axis="x", alpha=0.2, linestyle="--", color="#ccc")
    ax1.set_axisbelow(True)

    # ── Donut chart ──
    ax2 = fig.add_axes([0.52, 0.10, 0.46, 0.82])
    ax2.set_facecolor("#FFF8F0")
    _, _, autotexts = ax2.pie(
        counts, labels=None, colors=colors,
        autopct=lambda p: f"{p:.1f}%" if p > 3 else "",
        startangle=90, explode=[0.04]*len(counts),
        wedgeprops=dict(width=0.58, edgecolor="white", linewidth=3),
        pctdistance=0.78,
    )
    for at in autotexts:
        at.set_fontsize(9); at.set_fontweight("bold"); at.set_color("white")

    for r, fc in [(0.44,"#FFF8F0"),(0.36,"#FFE8EF"),(0.28,"#FFCCD5")]:
        ax2.add_patch(plt.Circle((0,0), r, fc=fc, zorder=10))
    ax2.text(0, 0.07, "CJ", ha="center", va="center", fontsize=22,
             fontweight="black", color="#C9184A", zorder=11)
    ax2.text(0, -0.14, "Smart", ha="center", va="center", fontsize=11,
             fontweight="bold", color="#FF4D6D", zorder=11)

    legend_handles = [
        mpatches.Patch(facecolor=colors[i], label=f"{labels[i]} ({counts[i]})")
        for i in range(len(cats))
    ]
    ax2.legend(handles=legend_handles, loc="lower center",
               bbox_to_anchor=(0.5,-0.18), ncol=2, fontsize=9,
               frameon=True, fancybox=True, framealpha=0.85, edgecolor="#ddd")
    ax2.set_title("Category Distribution", fontsize=13, fontweight="bold", color="#333", pad=10)

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor="#FFF8F0")
    plt.close(fig)
    buf.seek(0)
    return buf

# ─── EXCEL BUILDER ───
def style_header(ws, row, ncols, title, bg="C9184A"):
    thin = Side(style="thin", color="FFFFFF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    cell = ws.cell(row=row, column=1, value=title)
    cell.font = Font(bold=True, color="FFFFFF", size=13)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 28

def style_col_headers(ws, row, headers):
    fills = ["F8BBD0","FCE4EC","FFE0B2","E8F5E9","E3F2FD","EDE7F6","FFF8E1","F3E5F5"]
    thin = Side(style="thin", color="DDDDDD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col_idx, value=h)
        cell.font = Font(bold=True, color="555555", size=11)
        fill_color = fills[col_idx % len(fills)]
        cell.fill = PatternFill("solid", fgColor=fill_color)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    ws.row_dimensions[row].height = 22

def apply_row_style(ws, row, ncols, even=True):
    thin = Side(style="thin", color="EEEEEE")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    bg = "FFF5F5" if even else "FFFFFF"
    for col in range(1, ncols+1):
        cell = ws.cell(row=row, column=col)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.border = border
        cell.alignment = Alignment(vertical="center")

def build_excel(df, summary, branch_df, map_df, items_df=None):
    buf = io.BytesIO()

    # สร้างกราฟก่อน
    chart_buf = make_chart(summary)
    table_buf = make_top_branch_table(items_df) if items_df is not None else None

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="ข้อมูลทั้งหมด", index=False, startrow=2)
        summary.to_excel(writer, sheet_name="สรุปตามประเภท", index=False, startrow=2)
        branch_df.to_excel(writer, sheet_name="สรุปตามสาขา", index=False, startrow=2)
        map_df.to_excel(writer, sheet_name="mapping สินค้า", index=False, startrow=2)
        writer.book.create_sheet("กราฟสรุป")

    # เปิด workbook มาตกแต่ง
    buf.seek(0)
    wb = load_workbook(buf)

    cat_colors = {c: CATEGORIES.get(c,{"color":"9E9E9E"})["color"].replace("#","")
                  for c in summary["ประเภทสินค้า"].tolist()}

    # ── Sheet 1: ข้อมูลทั้งหมด ──
    ws1 = wb["ข้อมูลทั้งหมด"]
    style_header(ws1, 1, len(df.columns), "📋 ข้อมูลสินค้าทั้งหมด (วิเคราะห์ประเภทด้วย AI)")
    style_col_headers(ws1, 3, df.columns.tolist())
    for r in range(4, ws1.max_row+1):
        apply_row_style(ws1, r, len(df.columns), r%2==0)
        cat_cell = ws1.cell(row=r, column=df.columns.tolist().index("ประเภทสินค้า")+1)
        if cat_cell.value and cat_cell.value in cat_colors:
            col_hex = cat_colors[cat_cell.value]
            cat_cell.fill = PatternFill("solid", fgColor=col_hex+"33")
            cat_cell.font = Font(bold=True, color=col_hex)
    for col in ws1.columns:
        ws1.column_dimensions[get_column_letter(col[0].column)].width = 18
    ws1.freeze_panes = "A4"

    # ── Sheet 2: สรุปตามประเภท ──
    ws2 = wb["สรุปตามประเภท"]
    style_header(ws2, 1, len(summary.columns), "📊 สรุปจำนวนสินค้าตามประเภท")
    style_col_headers(ws2, 3, summary.columns.tolist())
    for r in range(4, ws2.max_row+1):
        apply_row_style(ws2, r, len(summary.columns), r%2==0)
        cat_cell = ws2.cell(row=r, column=1)
        if cat_cell.value and cat_cell.value in cat_colors:
            col_hex = cat_colors[cat_cell.value]
            for c in range(1, len(summary.columns)+1):
                cell = ws2.cell(row=r, column=c)
                cell.fill = PatternFill("solid", fgColor=col_hex+"22")
            cat_cell.font = Font(bold=True, color=col_hex)
    for col in ws2.columns:
        ws2.column_dimensions[get_column_letter(col[0].column)].width = 22

    # ── Sheet 3: สรุปตามสาขา ──
    ws3 = wb["สรุปตามสาขา"]
    style_header(ws3, 1, len(branch_df.columns), "🏪 สรุปตามสาขา")
    style_col_headers(ws3, 3, branch_df.columns.tolist())
    for r in range(4, ws3.max_row+1):
        apply_row_style(ws3, r, len(branch_df.columns), r%2==0)
    for col in ws3.columns:
        ws3.column_dimensions[get_column_letter(col[0].column)].width = 20

    # ── Sheet 4: mapping ──
    ws4 = wb["mapping สินค้า"]
    style_header(ws4, 1, 2, "🏷️ ตารางจับคู่สินค้า → ประเภท")
    style_col_headers(ws4, 3, ["ชื่อสินค้า","ประเภทสินค้า"])
    for r in range(4, ws4.max_row+1):
        apply_row_style(ws4, r, 2, r%2==0)
        cat_cell = ws4.cell(row=r, column=2)
        if cat_cell.value and cat_cell.value in cat_colors:
            col_hex = cat_colors[cat_cell.value]
            cat_cell.fill = PatternFill("solid", fgColor=col_hex+"33")
            cat_cell.font = Font(bold=True, color=col_hex)
    ws4.column_dimensions["A"].width = 35
    ws4.column_dimensions["B"].width = 22

    # ── Sheet 5: กราฟ ──
    ws5 = wb["กราฟสรุป"]
    ws5.sheet_view.showGridLines = False
    ws5.sheet_properties.tabColor = "C9184A"
    style_header(ws5, 1, 12, "📈 กราฟแสดงสัดส่วนประเภทสินค้า CJ Express")
    ws5.row_dimensions[1].height = 30
    chart_buf.seek(0)
    img = XLImage(chart_buf)
    img.anchor = "A3"
    img.width  = 1100
    img.height = 520
    ws5.add_image(img)
    if table_buf:
        table_buf.seek(0)
        tbl_img = XLImage(table_buf)
        tbl_img.anchor = "A35"
        tbl_img.width  = 1100
        tbl_img.height = 420
        ws5.add_image(tbl_img)
    for col_letter in "ABCDEFGHIJKL":
        ws5.column_dimensions[col_letter].width = 12

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# ─── HELPERS ───
def get_api_key():
    import os
    try:
        key = st.secrets["ANTHROPIC_API_KEY"]
        if key and str(key).strip(): return str(key).strip()
    except: pass
    try:
        key = st.secrets["secrets"]["ANTHROPIC_API_KEY"]
        if key and str(key).strip(): return str(key).strip()
    except: pass
    return os.environ.get("ANTHROPIC_API_KEY","").strip()

def claude_classify(products, api_key):
    payload = json.dumps({
        "model": "claude-sonnet-4-6", "max_tokens": 2000,
        "system": SYSTEM_PROMPT,
        "messages": [{"role":"user","content":"วิเคราะห์ประเภทสินค้าจากชื่อ:\n"+
                      "\n".join(f"- {p}" for p in products)}],
    }).encode()
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages", data=payload,
        headers={"Content-Type":"application/json","x-api-key":api_key,
                 "anthropic-version":"2023-06-01"},
    )
    with urllib.request.urlopen(req, timeout=60) as r:
        data = json.loads(r.read())
    raw = re.sub(r"```json|```","", data["content"][0]["text"]).strip()
    return json.loads(raw)

def classify_all(products, api_key):
    result, batch_size = {}, 20
    batches = [products[i:i+batch_size] for i in range(0,len(products),batch_size)]
    bar = st.progress(0)
    status = st.empty()
    for i, batch in enumerate(batches):
        status.markdown(f"🤖 วิเคราะห์ **{min((i+1)*batch_size,len(products))}/{len(products)}** รายการ...")
        try:    result.update(claude_classify(batch, api_key))
        except Exception as e:
            st.warning(f"Batch {i+1}: {e}")
            for p in batch: result[p] = "สินค้าเบ็ดเตล็ด"
        bar.progress((i+1)/len(batches))
    bar.empty(); status.empty()
    return result

def make_summary(items):
    grp = items.groupby("ประเภทสินค้า")
    df = pd.concat([grp["ชื่อสินค้า"].count(), grp["ยอดรวมสินค้า"].sum()], axis=1)
    df.columns = ["จำนวนรายการ","ยอดรวม"]
    return df.reset_index().sort_values("จำนวนรายการ", ascending=False)

def make_branch_summary(items):
    grp = items.groupby(["รหัสสาขา","ประเภทสินค้า"])
    df = pd.concat([grp["ชื่อสินค้า"].count(), grp["ยอดรวมสินค้า"].sum()], axis=1)
    df.columns = ["จำนวนรายการ","ยอดรวม"]
    return df.reset_index().sort_values(["รหัสสาขา","จำนวนรายการ"], ascending=[True,False])

def sec(title, icon=""):
    st.markdown(f'<div class="sec-title"><div class="dot"></div>{icon} {title}</div>',
                unsafe_allow_html=True)

# ─── SESSION STATE ───
for k, v in [("df",None),("cat_map",{}),("analyzed",False)]:
    if k not in st.session_state: st.session_state[k] = v

# ─── HERO ───
st.markdown("""
<div class="hero">
  <div class="hero-mascot">🛒</div>
  <div class="hero-title">CJ Smart Scan</div>
  <div class="hero-sub">📊 วิเคราะห์ประเภทสินค้าอัตโนมัติจากชื่อสินค้าด้วย AI</div>
</div>
""", unsafe_allow_html=True)

# ─── API KEY ───
api_key = get_api_key()
if not api_key:
    st.error("⚠️ ไม่พบ **ANTHROPIC_API_KEY**")
    with st.expander("🔍 Debug"):
        try: st.write("Keys:", list(st.secrets.keys()))
        except Exception as e: st.write("อ่าน secrets ไม่ได้:", e)
    st.stop()

# ─── LAYOUT ───
left, right = st.columns([1, 1.6], gap="large")

with left:
    sec("อัปโหลดไฟล์ Excel", "📂")
    uploaded = st.file_uploader("ลากไฟล์มาวางหรือกดเลือกไฟล์", type=["xlsx"],
                                 label_visibility="collapsed")
    if uploaded:
        try:
            df_raw = pd.read_excel(uploaded, sheet_name="ใบเสร็จ")
            drop_cols = [c for c in df_raw.columns
                         if "หมวด" in c or "ประเภท" in c or "category" in c.lower()]
            if drop_cols: df_raw = df_raw.drop(columns=drop_cols)
            st.session_state.df = df_raw
            st.session_state.analyzed = False
            st.session_state.cat_map = {}
            n_items  = int(df_raw["ชื่อสินค้า"].notna().sum())
            n_branch = int(df_raw["รหัสสาขา"].dropna().nunique())
            st.success(f"✅ โหลดสำเร็จ! พบ **{n_items}** รายการ จาก **{n_branch}** สาขา")
        except Exception as e:
            st.error(f"❌ {e}"); st.stop()

    if st.session_state.df is not None and not st.session_state.analyzed:
        if st.button("🤖  วิเคราะห์ประเภทสินค้า", type="primary", use_container_width=True):
            products = st.session_state.df["ชื่อสินค้า"].dropna().unique().tolist()
            st.session_state.cat_map = classify_all(products, api_key)
            st.session_state.analyzed = True
            st.rerun()

    if st.session_state.analyzed:
        sec("หมวดหมู่สินค้า", "🏷️")
        for name, meta in CATEGORIES.items():
            st.markdown(
                f'<div style="display:flex;align-items:center;gap:10px;padding:5px 0">' +
                f'<span style="font-size:1.3rem">{meta["icon"]}</span>' +
                f'<span style="background:{meta["color"]};color:white;padding:3px 12px;' +
                f'border-radius:99px;font-size:.8rem;font-weight:600">{name}</span>' +
                f'</div>', unsafe_allow_html=True)

with right:
    if st.session_state.analyzed and st.session_state.df is not None:
        df = st.session_state.df.copy()
        df["ประเภทสินค้า"] = df["ชื่อสินค้า"].map(st.session_state.cat_map)
        items = df[df["ชื่อสินค้า"].notna()].copy()

        summary_df = make_summary(items)
        branch_df  = make_branch_summary(items)
        map_df     = pd.DataFrame(
            sorted(st.session_state.cat_map.items(), key=lambda x: x[1]),
            columns=["ชื่อสินค้า","ประเภทสินค้า"])

        excel_bytes = build_excel(df, summary_df, branch_df, map_df, items_df=items)
        st.download_button(
            "⬇️  ดาวน์โหลด Excel (พร้อมกราฟ 5 sheets)",
            data=excel_bytes,
            file_name="CJ_product_classified.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.markdown("<br>", unsafe_allow_html=True)

        total_sales = float(items["ยอดรวมสินค้า"].sum())
        m1,m2,m3,m4 = st.columns(4)
        for col,icon,num,lbl in [
            (m1,"🏪",int(items["รหัสสาขา"].nunique()),"สาขา"),
            (m2,"📦",len(items),"รายการ"),
            (m3,"🏷️",int(items["ประเภทสินค้า"].nunique()),"หมวดหมู่"),
            (m4,"💰",f"฿{total_sales:,.0f}","ยอดรวม"),
        ]:
            col.markdown(f'<div class="mcard"><div class="icon">{icon}</div>' +
                         f'<div class="num">{num}</div><div class="lbl">{lbl}</div></div>',
                         unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        sec("สัดส่วนตามหมวดหมู่", "📊")
        max_cnt = int(summary_df["จำนวนรายการ"].max()) or 1
        for _, row in summary_df.iterrows():
            cat  = row["ประเภทสินค้า"]
            cnt  = int(row["จำนวนรายการ"])
            tot  = float(row["ยอดรวม"]) if pd.notna(row["ยอดรวม"]) else 0
            meta = CATEGORIES.get(cat, {"icon":"📦","color":"#9E9E9E"})
            pct  = cnt / max_cnt * 100
            st.markdown(
                f'<div class="cat-card">' +
                f'<div class="cat-icon">{meta["icon"]}</div>' +
                f'<div style="flex:1"><div class="cat-name">{cat}</div>' +
                f'<div class="cat-count">฿{tot:,.0f}</div>' +
                f'<div class="bar-wrap" style="margin-top:6px">' +
                f'<div class="bar-fill" style="width:{pct}%;background:{meta["color"]}"></div>' +
                f'</div></div><div class="cat-num">{cnt}</div></div>',
                unsafe_allow_html=True)

# ─── DETAIL TABLE ───
if st.session_state.analyzed and st.session_state.df is not None:
    df = st.session_state.df.copy()
    df["ประเภทสินค้า"] = df["ชื่อสินค้า"].map(st.session_state.cat_map)
    items = df[df["ชื่อสินค้า"].notna()].copy()

    st.markdown("<br>", unsafe_allow_html=True)
    sec("รายการสินค้าทั้งหมด", "📋")

    fa, fb, fc = st.columns(3)
    branch_ids = sorted(items["รหัสสาขา"].dropna().unique().astype(int).tolist())
    with fa:
        sel_branch = st.multiselect("🏪 สาขา", branch_ids, default=branch_ids,
                                    format_func=lambda x: f"สาขา {x}")
    with fb:
        sel_cat = st.multiselect("🏷️ ประเภท", list(CATEGORIES.keys()),
                                  default=list(CATEGORIES.keys()))
    with fc:
        search = st.text_input("🔍 ค้นหาชื่อสินค้า", placeholder="พิมพ์ชื่อสินค้า...")

    filtered = items[items["รหัสสาขา"].isin(sel_branch) &
                     items["ประเภทสินค้า"].isin(sel_cat)].copy()
    if search:
        filtered = filtered[filtered["ชื่อสินค้า"].str.contains(search, na=False, case=False)]

    show_cols = ["รหัสสาขา","วันที่","ชื่อสินค้า","ประเภทสินค้า","จำนวน","ราคาต่อหน่วย","ยอดรวมสินค้า"]
    def hl(val):
        meta = CATEGORIES.get(val, {"color":"#9E9E9E"})
        return f"background-color:{meta['color']}22;color:{meta['color']};font-weight:600"

    st.dataframe(
        filtered[show_cols].style
            .map(hl, subset=["ประเภทสินค้า"])
            .format({"จำนวน":"{:.0f}","ราคาต่อหน่วย":"฿{:,.2f}","ยอดรวมสินค้า":"฿{:,.2f}"}),
        use_container_width=True, hide_index=True, height=420)
    st.caption(f"แสดง **{len(filtered):,}** จาก **{len(items):,}** รายการ")

    st.markdown("<br>", unsafe_allow_html=True)
    sec("สรุปรายสาขา", "🏪")
    bdf2 = make_branch_summary(items)
    tabs = st.tabs([f"🏪 สาขา {b}" for b in branch_ids])
    for tab, b in zip(tabs, branch_ids):
        with tab:
            bd = bdf2[bdf2["รหัสสาขา"]==b]
            c1, c2 = st.columns([1.2,1])
            with c1:
                st.bar_chart(bd.set_index("ประเภทสินค้า")["จำนวนรายการ"],
                             color="#FF4D6D", height=260)
            with c2:
                st.dataframe(
                    bd[["ประเภทสินค้า","จำนวนรายการ","ยอดรวม"]]
                    .style.format({"ยอดรวม":"฿{:,.2f}"}),
                    use_container_width=True, hide_index=True, height=260)

st.markdown('<div class="footer">🛒 CJ Smart Scan · Powered by Claude AI · Made with ❤️</div>',
            unsafe_allow_html=True)
