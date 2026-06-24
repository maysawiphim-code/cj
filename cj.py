import streamlit as st
import pandas as pd
import json, re, io, urllib.request
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.series import DataPoint

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
    border-radius: 24px; padding: 32px 40px 28px; margin-bottom: 24px;
    position: relative; overflow: hidden;
    box-shadow: 0 8px 32px rgba(201,24,74,.25);
}
.hero-title {
    font-family:'Fredoka One','Sarabun',sans-serif !important;
    font-size:2.4rem; font-weight:900; color:white; margin:0;
    text-shadow:0 2px 8px rgba(0,0,0,.2);
}
.hero-sub { color:rgba(255,255,255,.88); font-size:1rem; margin:8px 0 0; }
.hero-mascot {
    position:absolute; right:40px; top:50%; transform:translateY(-50%);
    font-size:6rem; filter:drop-shadow(0 4px 12px rgba(0,0,0,.15));
}
.sec-title {
    display:flex; align-items:center; gap:10px;
    font-size:1.1rem; font-weight:700; color:#C9184A; margin:22px 0 12px;
}
.sec-title .dot { width:5px; height:22px; background:#FF4D6D; border-radius:3px; }
.mcard {
    background:white; border-radius:16px; padding:20px 14px; text-align:center;
    box-shadow:0 3px 12px rgba(0,0,0,.07); transition:transform .15s;
}
.mcard:hover { transform:translateY(-3px); }
.mcard .icon { font-size:2.2rem; line-height:1; margin-bottom:4px; }
.mcard .num  { font-size:1.8rem; font-weight:800; color:#C9184A; }
.mcard .lbl  { font-size:.78rem; color:#999; margin-top:4px; font-weight:500; }
.cat-card {
    background:#FFFBFC; border-radius:14px; padding:13px 18px; margin-bottom:8px;
    display:flex; align-items:center; gap:12px;
    box-shadow:0 1px 6px rgba(0,0,0,.05); border:1px solid #FFF0F3;
}
.cat-card .cat-icon { font-size:1.8rem; }
.cat-card .cat-name  { font-weight:700; font-size:.9rem; color:#333; }
.cat-card .cat-count { font-size:.78rem; color:#aaa; }
.cat-card .bar-wrap  { flex:1; background:#f0f0f0; border-radius:99px; height:7px; overflow:hidden; }
.cat-card .bar-fill  { height:100%; border-radius:99px; }
.cat-card .cat-num   { font-weight:800; font-size:1rem; color:#C9184A; min-width:32px; text-align:right; }
.compare-card {
    background:white; border-radius:16px; padding:18px 20px; margin-bottom:10px;
    box-shadow:0 2px 10px rgba(0,0,0,.06); border-left:4px solid #FF4D6D;
}
.stDownloadButton > button {
    background:linear-gradient(135deg,#FF6B6B,#C9184A) !important;
    color:white !important; border:none !important; border-radius:12px !important;
    padding:12px 24px !important; font-size:.95rem !important; font-weight:700 !important;
    box-shadow:0 4px 14px rgba(201,24,74,.3) !important;
}
.stButton > button[kind="primary"] {
    background:linear-gradient(135deg,#FF6B6B,#C9184A) !important;
    color:white !important; border:none !important; border-radius:12px !important;
    padding:12px 24px !important; font-size:.95rem !important; font-weight:700 !important;
}
.tab-content { padding: 8px 0; }
.stTabs [data-baseweb="tab-list"] {
    background:white; border-radius:12px; padding:4px; gap:4px;
    box-shadow:0 2px 8px rgba(0,0,0,.06);
}
.footer { text-align:center; color:#ccc; font-size:.78rem; padding:28px 0 12px; }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════
# CONSTANTS
# ═══════════════════════════════════════════
CATEGORIES = {
    "Bao Cafe":         {"icon":"☕","color":"#C2185B","hex":"C2185B"},
    "อาหารพร้อมทาน":   {"icon":"🍱","color":"#F57C00","hex":"F57C00"},
    "เครื่องดื่ม":      {"icon":"🥤","color":"#1976D2","hex":"1976D2"},
    "ขนมขบเคี้ยว":     {"icon":"🍿","color":"#7B1FA2","hex":"7B1FA2"},
    "ของใช้ส่วนตัว":   {"icon":"🧴","color":"#00897B","hex":"00897B"},
    "ของใช้ในบ้าน":    {"icon":"🏠","color":"#558B2F","hex":"558B2F"},
    "สินค้าเบ็ดเตล็ด": {"icon":"🔋","color":"#6D4C41","hex":"6D4C41"},
    "บริการและอื่นๆ":  {"icon":"📱","color":"#455A64","hex":"455A64"},
}
RANK_BG    = ["FFD700","C0C0C0","CD7F32","FF6B6B","FF9F43","48DBFB","1DD1A1","A29BFE","FD79A8","636E72"]
RANK_MEDAL = ["🥇","🥈","🥉","4","5","6","7","8","9","10"]

SYSTEM_PROMPT = """คุณคือระบบจำแนกประเภทสินค้าร้านสะดวกซื้อ CJ Express
วิเคราะห์จากชื่อสินค้าเท่านั้น
หมวดหมู่:
1.Bao Cafe→ขึ้นต้น Bao_ หรือ Bac_
2.อาหารพร้อมทาน→ปัง แซนด์วิช ไส้กรอก พุดดิ้ง ข้าวเกรียบ ไข่ แฮม
3.เครื่องดื่ม→น้ำดื่ม นม ชา กาแฟทั่วไป เบียร์ สุรา น้ำหวาน คอลลาเจน
4.ขนมขบเคี้ยว→มันฝรั่ง ลูกอม ช็อกโกแลต ไอศกรีม ถั่ว ขนมกรอบ
5.ของใช้ส่วนตัว→ยาสีฟัน แชมพู สบู่ ครีม โฟมล้างหน้า เครื่องสำอาง ยา
6.ของใช้ในบ้าน→ทิชชู ผงซักฟอก น้ำยาล้างจาน ถุงขยะ
7.สินค้าเบ็ดเตล็ด→ถ่านไฟฉาย ไฟแช็ก เครื่องเขียน เทปลบ
8.บริการและอื่นๆ→เติมเงิน จ่ายบิล ส่วนลด แก้วลูกค้า
ตอบ JSON เท่านั้น: {"ชื่อสินค้า":"หมวดหมู่"}"""

RULE_KEYWORDS = {
    "Bao Cafe":        ["bao_","bao ","bac_","bac ","เอสเปรสโซ","ลาเต้","อเมริกาโน","มัทฉะ","อาราบีก้า","ชาเขียวนมสด"],
    "อาหารพร้อมทาน":  ["ปัง","แซนด์วิช","ไส้กรอก","พุดดิ้ง","ข้าวเกรียบ","ผงชูรส","ไข่","แฮม","มินิบัน","เบเกอรี"],
    "เครื่องดื่ม":     ["น้ํา","น้ำ","นม","เบียร์","สุรา","โคล่า","โซดา","คอลลาเจน","โพรไบโอ","ชูกำลัง","เฮล","วิตซี","ซีวิท"],
    "ขนมขบเคี้ยว":    ["มันฝรั่ง","ลูกอม","ช็อก","ไอศกรีม","ถั่ว","เมล็ด","อัลมอนด์","ข้าวโพด","ทวิสโก้","คอนเน่","บันบัน","ซาซ่า","กินดะ","โรซี่","โคอาล่า","ตะวันขนม","ฟู่อารี่","เทสโต","เลยร้อย","เลยร็อค"],
    "ของใช้ส่วนตัว":  ["ยาสีฟัน","แชมพู","สบู่","ครีม","ผ้าอนามัย","โฟม","รองพื้น","แปรง","เดนทีน","บีโอเร","การ์นิเย่","สกินแสบ","เจเล","แพนดิน","สุภาภรณ์","เอล สมูธ","mbl","bigsml"],
    "ของใช้ในบ้าน":   ["ทิชชู","ผงซัก","น้ำยา","ถุงขยะ","ดาวน์นี่","วิกซอล","ปรับผ้า"],
    "สินค้าเบ็ดเตล็ด":["ถ่าน","ไฟแช็ก","ยากันยุง","เครื่องเขียน","เทปลบ","พานาโซนิค","ดราช่าง","เรนเจอร์"],
    "บริการและอื่นๆ": ["เติมเงิน","จ่ายบิล","ซิมการ์ด","บัตรเติม","ส่วนลด","แก้วลูกค้า","00"],
}

# ═══════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════
def get_api_key():
    import os
    try:
        k = st.secrets["GEMINI_API_KEY"]
        if k and str(k).strip(): return str(k).strip()
    except: pass
    return os.environ.get("GEMINI_API_KEY","").strip()

def rule_classify(name):
    n = str(name).lower()
    for cat, kws in RULE_KEYWORDS.items():
        if any(k.lower() in n for k in kws): return cat
    return ""

def classify_rule_only(products):
    return {p: rule_classify(p) or "สินค้าเบ็ดเตล็ด" for p in products}

def gemini_classify(products, api_key):
    product_list = "\n".join(f"- {p}" for p in products)
    prompt = SYSTEM_PROMPT + "\n\nวิเคราะห์:\n" + product_list
    payload = json.dumps({
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.1, "maxOutputTokens": 2000},
    }, ensure_ascii=False).encode("utf-8")
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={api_key}"
    req = urllib.request.Request(url, data=payload,
                                  headers={"Content-Type":"application/json; charset=utf-8"})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            data = json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        raise Exception(f"API {e.code}: {e.read().decode()[:200]}")
    raw = re.sub(r"```json|```","", data["candidates"][0]["content"]["parts"][0]["text"]).strip()
    return json.loads(raw)

def classify_with_ai(products, api_key, current_map):
    need_ai = [p for p in products if current_map.get(p)=="สินค้าเบ็ดเตล็ด" and not rule_classify(p)]
    if not need_ai:
        return current_map
    result = dict(current_map)
    batches = [need_ai[i:i+30] for i in range(0,len(need_ai),30)]
    bar = st.progress(0); status = st.empty()
    for i, batch in enumerate(batches):
        status.markdown(f"🤖 AI วิเคราะห์ **{i*30+len(batch)}/{len(need_ai)}** รายการ...")
        try:
            result.update(gemini_classify(batch, api_key))
        except Exception as e:
            err = str(e)
            if "429" in err or "quota" in err.lower():
                st.warning(f"⚠️ Rate limit — {len(need_ai)-i*30} รายการที่เหลือใช้ rule-based แทน")
                break
            else:
                st.warning(f"Batch {i+1}: {err[:120]}")
        bar.progress((i+1)/len(batches))
    bar.empty(); status.empty()
    return result

def load_excel(file):
    df = pd.read_excel(file, sheet_name="ใบเสร็จ")
    drop = [c for c in df.columns if "หมวด" in c or "ประเภท" in c or "category" in c.lower()]
    if drop: df = df.drop(columns=drop)
    df["วันที่_dt"] = pd.to_datetime(df["วันที่"], format="%d/%m/%Y", errors="coerce")
    df["เดือน"]     = df["วันที่_dt"].dt.strftime("%b %Y")
    df["เดือน_sort"]= df["วันที่_dt"].dt.to_period("M").astype(str)
    return df

def make_summary(items):
    grp = items.groupby("ประเภทสินค้า")
    df  = pd.concat([grp["ชื่อสินค้า"].count(), grp["ยอดรวมสินค้า"].sum()], axis=1)
    df.columns = ["จำนวนรายการ","ยอดรวม"]
    return df.reset_index().sort_values("จำนวนรายการ", ascending=False)

def make_branch_summary(items):
    grp = items.groupby(["รหัสสาขา","ประเภทสินค้า"])
    df  = pd.concat([grp["ชื่อสินค้า"].count(), grp["ยอดรวมสินค้า"].sum()], axis=1)
    df.columns = ["จำนวนรายการ","ยอดรวม"]
    return df.reset_index().sort_values(["รหัสสาขา","จำนวนรายการ"], ascending=[True,False])

def make_monthly_top(items):
    df = items.copy()
    # สร้างคอลัมน์เดือนถ้ายังไม่มี
    if "เดือน_sort" not in df.columns:
        df["วันที่_dt"]  = pd.to_datetime(df["วันที่"], format="%d/%m/%Y", errors="coerce")
        df["เดือน_sort"] = df["วันที่_dt"].dt.to_period("M").astype(str)
        df["เดือน"]      = df["วันที่_dt"].dt.strftime("%b %Y")
    cat_grp = df.groupby(["เดือน_sort","รหัสสาขา","ประเภทสินค้า"])["ชื่อสินค้า"].count().reset_index()
    cat_grp = cat_grp.sort_values("ชื่อสินค้า",ascending=False).drop_duplicates(["เดือน_sort","รหัสสาขา"])
    cat_grp = cat_grp.rename(columns={"ชื่อสินค้า":"_c","ประเภทสินค้า":"top_cat"})
    grp = df.groupby(["เดือน_sort","เดือน","รหัสสาขา"])
    m   = pd.concat([grp["ชื่อสินค้า"].count(), grp["ยอดรวมสินค้า"].sum()], axis=1).reset_index()
    m.columns = ["เดือน_sort","เดือน","รหัสสาขา","จำนวนรายการ","ยอดรวม"]
    m = m.merge(cat_grp[["เดือน_sort","รหัสสาขา","top_cat"]], on=["เดือน_sort","รหัสสาขา"], how="left")
    return m.sort_values(["เดือน_sort","ยอดรวม"], ascending=[True,False])

def sec(title, icon=""):
    st.markdown(f'<div class="sec-title"><div class="dot"></div>{icon} {title}</div>', unsafe_allow_html=True)

def cat_badge(cat, size="sm"):
    meta = CATEGORIES.get(cat, {"color":"#9E9E9E","icon":"📦"})
    fs = ".75rem" if size=="sm" else ".9rem"
    return (f'<span style="background:{meta["color"]}18;color:{meta["color"]};'
            f'padding:3px 10px;border-radius:99px;font-size:{fs};font-weight:700;'
            f'border:1px solid {meta["color"]}44">{meta["icon"]} {cat}</span>')

# ═══════════════════════════════════════════
# EXCEL BUILDER
# ═══════════════════════════════════════════
def thin_border():
    s = Side(style="thin", color="E8E8E8")
    return Border(left=s, right=s, top=s, bottom=s)

def write_header(ws, row, ncols, title, bg="C9184A"):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=title)
    c.font = Font(bold=True, color="FFFFFF", size=13)
    c.fill = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 28

def write_col_headers(ws, row, headers, bg="FCE4EC"):
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=j, value=h)
        c.font = Font(bold=True, color="555555", size=10)
        c.fill = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border()
    ws.row_dimensions[row].height = 20

def stripe_row(ws, row, ncols, even=True):
    bg = "FFF8FA" if even else "FFFFFF"   # อ่อนลงมาก
    for col in range(1, ncols+1):
        c = ws.cell(row=row, column=col)
        if not c.fill or c.fill.fgColor.rgb in ("00000000","FFFFFFFF","FFF8FA","FFFFFF"):
            c.fill = PatternFill("solid", fgColor=bg)
        c.border = thin_border()
        c.alignment = Alignment(vertical="center")

def build_excel(df, summary, branch_df, map_df, items, df_prev=None, summary_prev=None):
    buf = io.BytesIO()
    sheets = {"ข้อมูลทั้งหมด":df, "สรุปตามประเภท":summary,
              "สรุปตามสาขา":branch_df, "mapping สินค้า":map_df}
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sname, sdf in sheets.items():
            sdf.to_excel(writer, sheet_name=sname, index=False, startrow=2)
        writer.book.create_sheet("กราฟ & Top สาขา")
        if df_prev is not None:
            writer.book.create_sheet("เปรียบเทียบเดือน")
    buf.seek(0)
    wb = load_workbook(buf)
    cat_hex = {c: CATEGORIES.get(c,{"hex":"9E9E9E"})["hex"] for c in summary["ประเภทสินค้า"]}

    # ── Sheet 1 ──
    ws1 = wb["ข้อมูลทั้งหมด"]
    write_header(ws1, 1, len(df.columns), "📋 ข้อมูลสินค้าทั้งหมด")
    write_col_headers(ws1, 3, df.columns.tolist())
    cat_idx = df.columns.tolist().index("ประเภทสินค้า")+1 if "ประเภทสินค้า" in df.columns else None
    for r in range(4, ws1.max_row+1):
        stripe_row(ws1, r, len(df.columns), r%2==0)
        if cat_idx:
            cv = ws1.cell(row=r, column=cat_idx).value
            if cv and cv in cat_hex:
                h = cat_hex[cv]
                ws1.cell(row=r,column=cat_idx).fill = PatternFill("solid", fgColor=h+"22")
                ws1.cell(row=r,column=cat_idx).font = Font(bold=True, color=h)
    for col in ws1.columns:
        ws1.column_dimensions[get_column_letter(col[0].column)].width = 18
    ws1.freeze_panes = "A4"

    # ── Sheet 2: สรุปตามประเภท ──
    ws2 = wb["สรุปตามประเภท"]
    write_header(ws2, 1, len(summary.columns), "📊 สรุปตามประเภทสินค้า")
    write_col_headers(ws2, 3, summary.columns.tolist())
    for r in range(4, ws2.max_row+1):
        cv = ws2.cell(row=r, column=1).value
        h  = cat_hex.get(cv, "F5F5F5")
        bg = h+"15"   # อ่อนมาก
        for c in range(1, len(summary.columns)+1):
            cell = ws2.cell(row=r, column=c)
            cell.fill   = PatternFill("solid", fgColor=bg)
            cell.border = thin_border()
            cell.alignment = Alignment(vertical="center", horizontal="center")
        ws2.cell(row=r,column=1).font = Font(bold=True, color=h if h!="F5F5F5" else "333333")
    for col in ws2.columns:
        ws2.column_dimensions[get_column_letter(col[0].column)].width = 22

    # ── Sheet 3 ──
    ws3 = wb["สรุปตามสาขา"]
    write_header(ws3, 1, len(branch_df.columns), "🏪 สรุปตามสาขา")
    write_col_headers(ws3, 3, branch_df.columns.tolist())
    for r in range(4, ws3.max_row+1):
        stripe_row(ws3, r, len(branch_df.columns), r%2==0)
    for col in ws3.columns:
        ws3.column_dimensions[get_column_letter(col[0].column)].width = 20

    # ── Sheet 4 ──
    ws4 = wb["mapping สินค้า"]
    write_header(ws4, 1, 2, "🏷️ mapping สินค้า → ประเภท")
    write_col_headers(ws4, 3, ["ชื่อสินค้า","ประเภทสินค้า"])
    for r in range(4, ws4.max_row+1):
        stripe_row(ws4, r, 2, r%2==0)
        cv = ws4.cell(row=r, column=2).value
        if cv and cv in cat_hex:
            h = cat_hex[cv]
            ws4.cell(row=r,column=2).fill = PatternFill("solid", fgColor=h+"22")
            ws4.cell(row=r,column=2).font = Font(bold=True, color=h)
    ws4.column_dimensions["A"].width = 35
    ws4.column_dimensions["B"].width = 22

    # ── Sheet 5: กราฟ + Top สาขา ──
    ws5 = wb["กราฟ & Top สาขา"]
    ws5.sheet_view.showGridLines = False
    ws5.sheet_properties.tabColor = "C9184A"
    for i in range(1,18): ws5.column_dimensions[get_column_letter(i)].width = 14
    write_header(ws5, 1, 16, "📈 กราฟ & Top 10 สาขาขายดี")

    # data for charts (col Q-R)
    ws5.cell(row=3,column=17,value="ประเภท")
    ws5.cell(row=3,column=18,value="จำนวน")
    for i,(_, row) in enumerate(summary.iterrows(),1):
        ws5.cell(row=3+i,column=17,value=row["ประเภทสินค้า"])
        ws5.cell(row=3+i,column=18,value=int(row["จำนวนรายการ"]))
    n = len(summary)

    bar = BarChart(); bar.type="bar"; bar.grouping="clustered"
    bar.title="จำนวนรายการตามประเภท"; bar.style=10; bar.width=18; bar.height=12
    bar.add_data(Reference(ws5,min_col=18,min_row=3,max_row=3+n),titles_from_data=True)
    bar.set_categories(Reference(ws5,min_col=17,min_row=4,max_row=3+n))
    bar.series[0].graphicalProperties.solidFill = "FF4D6D"
    bar.series[0].graphicalProperties.line.solidFill = "C9184A"
    ws5.add_chart(bar,"A3")

    pie = PieChart(); pie.title="สัดส่วนประเภทสินค้า"; pie.style=10; pie.width=14; pie.height=12
    pie.add_data(Reference(ws5,min_col=18,min_row=3,max_row=3+n),titles_from_data=True)
    pie.set_categories(Reference(ws5,min_col=17,min_row=4,max_row=3+n))
    ws5.add_chart(pie,"L3")

    # Top 10 table
    monthly = make_monthly_top(items)
    cr = 24
    write_header(ws5, cr, 6, "🏆 Top 10 สาขาขายดีรายเดือน", bg="880E4F")
    ws5.row_dimensions[cr].height = 28; cr+=1

    for month_key in sorted(monthly["เดือน_sort"].unique()):
        mdata = monthly[monthly["เดือน_sort"]==month_key].head(10)
        month_lbl = mdata["เดือน"].iloc[0]
        ws5.merge_cells(start_row=cr,start_column=1,end_row=cr,end_column=6)
        mc = ws5.cell(row=cr,column=1,value=f"📅  {month_lbl}")
        mc.font=Font(bold=True,color="FFFFFF",size=11)
        mc.fill=PatternFill("solid",fgColor="C9184A")
        mc.alignment=Alignment(horizontal="left",vertical="center",indent=1)
        ws5.row_dimensions[cr].height=22; cr+=1

        write_col_headers(ws5,cr,["อันดับ","รหัสสาขา","เดือน","ยอดรวม","จำนวน","ประเภทนิยม"],"FFE0B2")
        ws5.row_dimensions[cr].height=18; cr+=1

        for i,(_, row) in enumerate(mdata.iterrows()):
            rh = RANK_BG[i] if i<len(RANK_BG) else "AAAAAA"
            cat = str(row.get("top_cat",""))
            ch  = CATEGORIES.get(cat,{"hex":"9E9E9E"})["hex"]
            bg  = "FFF8FA" if i%2==0 else "FFFFFF"
            vals=[RANK_MEDAL[i] if i<len(RANK_MEDAL) else str(i+1),
                  str(int(row["รหัสสาขา"])), str(row["เดือน"]),
                  f'{row["ยอดรวม"]:,.2f}', str(int(row["จำนวนรายการ"])), cat]
            for j,val in enumerate(vals,1):
                c=ws5.cell(row=cr,column=j,value=val)
                c.alignment=Alignment(horizontal="center",vertical="center")
                c.border=thin_border()
                if j==1: c.fill=PatternFill("solid",fgColor=rh); c.font=Font(bold=True,color="FFFFFF",size=10)
                elif j==6: c.fill=PatternFill("solid",fgColor=ch+"22"); c.font=Font(bold=True,color=ch,size=9)
                else: c.fill=PatternFill("solid",fgColor=bg); c.font=Font(color="333333",size=9)
            ws5.row_dimensions[cr].height=18; cr+=1
        cr+=1

    # ── Sheet 6: เปรียบเทียบ (ถ้ามี) ──
    if df_prev is not None and summary_prev is not None:
        ws6 = wb["เปรียบเทียบเดือน"]
        ws6.sheet_view.showGridLines = False
        ws6.sheet_properties.tabColor = "1976D2"
        for i in range(1,14): ws6.column_dimensions[get_column_letter(i)].width = 18
        write_header(ws6,1,10,"📊 เปรียบเทียบยอดขายระหว่างสองเดือน","1565C0")
        ws6.row_dimensions[1].height=28

        # merge summary
        merged = summary.rename(columns={"จำนวนรายการ":"จำนวน_ปัจจุบัน","ยอดรวม":"ยอด_ปัจจุบัน"}).merge(
            summary_prev.rename(columns={"จำนวนรายการ":"จำนวน_เดือนก่อน","ยอดรวม":"ยอด_เดือนก่อน"}),
            on="ประเภทสินค้า", how="outer").fillna(0)
        merged["เปลี่ยนแปลง%"] = merged.apply(
            lambda r: f'+{((r["ยอด_ปัจจุบัน"]-r["ยอด_เดือนก่อน"])/r["ยอด_เดือนก่อน"]*100):.1f}%'
            if r["ยอด_เดือนก่อน"]>0
            else "ใหม่" if r["ยอด_ปัจจุบัน"]>0 else "-", axis=1)

        hdrs=["ประเภทสินค้า","จำนวน_เดือนก่อน","ยอด_เดือนก่อน","จำนวน_ปัจจุบัน","ยอด_ปัจจุบัน","เปลี่ยนแปลง%"]
        write_col_headers(ws6,3,hdrs,"BBDEFB")
        for r_idx,(_, row) in enumerate(merged[hdrs].iterrows(),4):
            cat = row["ประเภทสินค้า"]
            h   = CATEGORIES.get(cat,{"hex":"9E9E9E"})["hex"]
            pct = str(row["เปลี่ยนแปลง%"])
            up  = pct.startswith("+")
            for j,val in enumerate(hdrs,1):
                c = ws6.cell(row=r_idx,column=j,value=row[val])
                c.border=thin_border()
                c.alignment=Alignment(horizontal="center",vertical="center")
                if j==1:
                    c.fill=PatternFill("solid",fgColor=h+"18")
                    c.font=Font(bold=True,color=h)
                elif j==6:
                    c.fill=PatternFill("solid",fgColor="E8F5E9" if up else "FFEBEE")
                    c.font=Font(bold=True,color="2E7D32" if up else "C62828")
                else:
                    c.fill=PatternFill("solid",fgColor="F8F8F8" if r_idx%2==0 else "FFFFFF")
                    c.font=Font(color="333333",size=10)
            ws6.row_dimensions[r_idx].height=20

        # chart เปรียบเทียบ
        ws6.cell(row=3,column=8,value="ประเภท")
        ws6.cell(row=3,column=9,value="เดือนก่อน")
        ws6.cell(row=3,column=10,value="ปัจจุบัน")
        for i,(_, row) in enumerate(merged.iterrows(),1):
            ws6.cell(row=3+i,column=8,value=row["ประเภทสินค้า"])
            ws6.cell(row=3+i,column=9,value=float(row["ยอด_เดือนก่อน"]))
            ws6.cell(row=3+i,column=10,value=float(row["ยอด_ปัจจุบัน"]))
        nm = len(merged)
        cbar = BarChart(); cbar.type="col"; cbar.grouping="clustered"
        cbar.title="เปรียบเทียบยอดขายตามประเภท"; cbar.style=10; cbar.width=20; cbar.height=14
        cbar.add_data(Reference(ws6,min_col=9,min_row=3,max_col=10,max_row=3+nm),titles_from_data=True)
        cbar.set_categories(Reference(ws6,min_col=8,min_row=4,max_row=3+nm))
        cbar.series[0].graphicalProperties.solidFill="90CAF9"
        cbar.series[1].graphicalProperties.solidFill="EF5350"
        ws6.add_chart(cbar,f"A{r_idx+3}")

    out = io.BytesIO(); wb.save(out); return out.getvalue()

# ═══════════════════════════════════════════
# SESSION STATE
# ═══════════════════════════════════════════
for k,v in [("df",None),("cat_map",{}),("analyzed",False),("_file_id",""),
             ("df_prev",None),("cat_map_prev",{}),("analyzed_prev",False),("_file_id_prev",""),
             ("ai_done",False)]:
    if k not in st.session_state: st.session_state[k] = v

api_key = get_api_key()

# ═══════════════════════════════════════════
# HERO
# ═══════════════════════════════════════════
st.markdown("""
<div class="hero">
  <div class="hero-mascot">🛒</div>
  <div class="hero-title">CJ Smart Scan</div>
  <div class="hero-sub">📊 วิเคราะห์ & เปรียบเทียบประเภทสินค้ารายเดือน ด้วย AI</div>
</div>
""", unsafe_allow_html=True)

if not api_key:
    st.error("⚠️ ไม่พบ **GEMINI_API_KEY** — ตั้งค่าใน Streamlit Secrets")
    st.code('GEMINI_API_KEY = "AIza..."', language="toml")
    st.stop()

# ═══════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════
tab1, tab2 = st.tabs(["📦 เดือนปัจจุบัน", "📊 เปรียบเทียบกับเดือนก่อน"])

# ══════════════════════════════════
# TAB 1 — เดือนปัจจุบัน
# ══════════════════════════════════
with tab1:
    left, right = st.columns([1, 1.6], gap="large")

    with left:
        sec("อัปโหลดไฟล์ Excel เดือนปัจจุบัน", "📂")
        uploaded = st.file_uploader("เลือกไฟล์ Excel", type=["xlsx"],
                                     key="up1", label_visibility="collapsed")
        if uploaded is not None:
            fid = f"{uploaded.name}_{uploaded.size}"
            if st.session_state._file_id != fid:
                try:
                    df_raw = load_excel(uploaded)
                    st.session_state.df        = df_raw
                    st.session_state.analyzed  = False
                    st.session_state.cat_map   = {}
                    st.session_state._file_id  = fid
                except Exception as e:
                    st.error(f"❌ {e}"); st.stop()

        if st.session_state.df is not None:
            n = int(st.session_state.df["ชื่อสินค้า"].notna().sum())
            b = int(st.session_state.df["รหัสสาขา"].dropna().nunique())
            st.success(f"✅ พบ **{n}** รายการ จาก **{b}** สาขา")

        if st.session_state.df is not None and not st.session_state.analyzed:
            if st.button("⚡ จำแนกสินค้าทันที", type="primary", use_container_width=True, key="btn_classify"):
                prods = st.session_state.df["ชื่อสินค้า"].dropna().unique().tolist()
                st.session_state.cat_map  = classify_rule_only(prods)
                st.session_state.analyzed = True
                st.rerun()

        if st.session_state.analyzed and st.session_state.df is not None:
            prods   = st.session_state.df["ชื่อสินค้า"].dropna().unique().tolist()
            need_ai = [p for p in prods if st.session_state.cat_map.get(p)=="สินค้าเบ็ดเตล็ด" and not rule_classify(p)]
            if need_ai:
                st.info(f"มี **{len(need_ai)}** รายการไม่แน่ใจ")
                if st.button("🤖 วิเคราะห์เพิ่มด้วย Gemini AI", use_container_width=True, key="btn_ai"):
                    new_map = classify_with_ai(prods, api_key, st.session_state.cat_map)
                    st.session_state.cat_map = new_map
                    still = [p for p in prods if new_map.get(p)=="สินค้าเบ็ดเตล็ด" and not rule_classify(p)]
                    st.session_state.ai_done = len(still)==0
                    st.rerun()
            else:
                st.success("✅ วิเคราะห์ครบทุกรายการแล้ว 🎉")

            if st.session_state.ai_done:
                st.success("🎉 AI วิเคราะห์ครบทุกรายการแล้ว!")
                st.session_state.ai_done = False

            sec("หมวดหมู่", "🏷️")
            for name, meta in CATEGORIES.items():
                st.markdown(
                    f'<div style="display:flex;align-items:center;gap:10px;padding:4px 0">'
                    f'<span style="font-size:1.2rem">{meta["icon"]}</span>'
                    f'<span style="background:{meta["color"]}18;color:{meta["color"]};padding:3px 12px;'
                    f'border-radius:99px;font-size:.78rem;font-weight:700;border:1px solid {meta["color"]}44">'
                    f'{name}</span></div>', unsafe_allow_html=True)

    with right:
        if st.session_state.analyzed and st.session_state.df is not None:
            df = st.session_state.df.copy()
            df["ประเภทสินค้า"] = df["ชื่อสินค้า"].map(st.session_state.cat_map)
            items = df[df["ชื่อสินค้า"].notna()].copy()
            summary_df = make_summary(items)
            branch_df  = make_branch_summary(items)
            map_df     = pd.DataFrame(sorted(st.session_state.cat_map.items(),key=lambda x:x[1]),
                                      columns=["ชื่อสินค้า","ประเภทสินค้า"])

            # download
            has_prev = st.session_state.df_prev is not None and st.session_state.analyzed_prev
            sp = None
            if has_prev:
                df_p = st.session_state.df_prev.copy()
                df_p["ประเภทสินค้า"] = df_p["ชื่อสินค้า"].map(st.session_state.cat_map_prev)
                ip   = df_p[df_p["ชื่อสินค้า"].notna()].copy()
                sp   = make_summary(ip)

            excel_bytes = build_excel(df, summary_df, branch_df, map_df, items,
                                      df_prev=st.session_state.df_prev if has_prev else None,
                                      summary_prev=sp)
            lbl = "⬇️ ดาวน์โหลด Excel (พร้อมเปรียบเทียบ)" if has_prev else "⬇️ ดาวน์โหลด Excel"
            st.download_button(lbl, data=excel_bytes, file_name="CJ_classified.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)
            st.markdown("<br>", unsafe_allow_html=True)

            # metrics
            m1,m2,m3,m4 = st.columns(4)
            for col,icon,num,lbl in [
                (m1,"🏪",int(items["รหัสสาขา"].nunique()),"สาขา"),
                (m2,"📦",len(items),"รายการ"),
                (m3,"🏷️",int(items["ประเภทสินค้า"].nunique()),"หมวดหมู่"),
                (m4,"💰",f'฿{items["ยอดรวมสินค้า"].sum():,.0f}',"ยอดรวม"),
            ]:
                col.markdown(f'<div class="mcard"><div class="icon">{icon}</div>'
                             f'<div class="num">{num}</div><div class="lbl">{lbl}</div></div>',
                             unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)

            # category bars
            sec("สัดส่วนตามหมวดหมู่", "📊")
            max_cnt = int(summary_df["จำนวนรายการ"].max()) or 1
            for _, row in summary_df.iterrows():
                cat  = row["ประเภทสินค้า"]; cnt=int(row["จำนวนรายการ"])
                tot  = float(row["ยอดรวม"]) if pd.notna(row["ยอดรวม"]) else 0
                meta = CATEGORIES.get(cat,{"icon":"📦","color":"#9E9E9E"})
                pct  = cnt/max_cnt*100
                st.markdown(
                    f'<div class="cat-card">'
                    f'<div class="cat-icon">{meta["icon"]}</div>'
                    f'<div style="flex:1"><div class="cat-name">{cat}</div>'
                    f'<div class="cat-count">฿{tot:,.0f}</div>'
                    f'<div class="bar-wrap" style="margin-top:5px">'
                    f'<div class="bar-fill" style="width:{pct}%;background:{meta["color"]}"></div>'
                    f'</div></div><div class="cat-num">{cnt}</div></div>',
                    unsafe_allow_html=True)

            # detail table
            st.markdown("<br>", unsafe_allow_html=True)
            sec("รายการสินค้า", "📋")
            fa,fb,fc = st.columns(3)
            bids = sorted(items["รหัสสาขา"].dropna().unique().astype(int).tolist())
            with fa: sel_b = st.multiselect("🏪 สาขา",bids,default=bids,key="sb1",format_func=lambda x:f"สาขา {x}")
            with fb: sel_c = st.multiselect("🏷️ ประเภท",list(CATEGORIES.keys()),default=list(CATEGORIES.keys()),key="sc1")
            with fc: srch  = st.text_input("🔍 ค้นหา",placeholder="ชื่อสินค้า...",key="sr1")
            filt = items[items["รหัสสาขา"].isin(sel_b)&items["ประเภทสินค้า"].isin(sel_c)].copy()
            if srch: filt = filt[filt["ชื่อสินค้า"].str.contains(srch,na=False,case=False)]
            cols = ["รหัสสาขา","วันที่","ชื่อสินค้า","ประเภทสินค้า","จำนวน","ราคาต่อหน่วย","ยอดรวมสินค้า"]
            def hl(v): m=CATEGORIES.get(v,{"color":"#9E9E9E"}); return f"background:{m['color']}15;color:{m['color']};font-weight:700"
            st.dataframe(filt[cols].style.map(hl,subset=["ประเภทสินค้า"])
                         .format({"จำนวน":"{:.0f}","ราคาต่อหน่วย":"฿{:,.2f}","ยอดรวมสินค้า":"฿{:,.2f}"}),
                         use_container_width=True, hide_index=True, height=400)
            st.caption(f"แสดง **{len(filt):,}** จาก **{len(items):,}** รายการ")

            # branch tabs
            st.markdown("<br>", unsafe_allow_html=True)
            sec("สรุปรายสาขา","🏪")
            bdf2 = make_branch_summary(items)
            tabs_b = st.tabs([f"🏪 {b}" for b in bids])
            for tab,b in zip(tabs_b,bids):
                with tab:
                    bd=bdf2[bdf2["รหัสสาขา"]==b]
                    c1,c2=st.columns([1.2,1])
                    with c1: st.bar_chart(bd.set_index("ประเภทสินค้า")["จำนวนรายการ"],color="#FF4D6D",height=240)
                    with c2: st.dataframe(bd[["ประเภทสินค้า","จำนวนรายการ","ยอดรวม"]].style.format({"ยอดรวม":"฿{:,.2f}"}),use_container_width=True,hide_index=True,height=240)

# ══════════════════════════════════
# TAB 2 — เปรียบเทียบเดือนก่อน
# ══════════════════════════════════
with tab2:
    st.markdown("### 📊 เปรียบเทียบกับเดือนก่อน")

    col_up, col_info = st.columns([1,1.6], gap="large")

    with col_up:
        sec("อัปโหลดไฟล์เดือนก่อน","📂")
        uploaded_prev = st.file_uploader("เลือกไฟล์ Excel เดือนก่อน", type=["xlsx"],
                                          key="up2", label_visibility="collapsed")
        if uploaded_prev is not None:
            fid2 = f"{uploaded_prev.name}_{uploaded_prev.size}"
            if st.session_state._file_id_prev != fid2:
                try:
                    df_p = load_excel(uploaded_prev)
                    st.session_state.df_prev        = df_p
                    st.session_state.analyzed_prev  = False
                    st.session_state.cat_map_prev   = {}
                    st.session_state._file_id_prev  = fid2
                except Exception as e:
                    st.error(f"❌ {e}"); st.stop()

        if st.session_state.df_prev is not None:
            n2 = int(st.session_state.df_prev["ชื่อสินค้า"].notna().sum())
            b2 = int(st.session_state.df_prev["รหัสสาขา"].dropna().nunique())
            st.success(f"✅ พบ **{n2}** รายการ จาก **{b2}** สาขา")

        if st.session_state.df_prev is not None and not st.session_state.analyzed_prev:
            if st.button("⚡ จำแนกไฟล์เดือนก่อน", type="primary", use_container_width=True, key="btn_prev"):
                prods2 = st.session_state.df_prev["ชื่อสินค้า"].dropna().unique().tolist()
                st.session_state.cat_map_prev   = classify_rule_only(prods2)
                st.session_state.analyzed_prev  = True
                st.rerun()

        if not st.session_state.analyzed:
            st.info("💡 กรุณาจำแนกไฟล์เดือนปัจจุบันในแท็บแรกก่อน")
        if not st.session_state.analyzed_prev and st.session_state.df_prev is not None:
            st.info("💡 กดจำแนกไฟล์เดือนก่อนด้านบน")

    with col_info:
        if st.session_state.analyzed and st.session_state.analyzed_prev:
            # เตรียมข้อมูล
            df_cur = st.session_state.df.copy()
            df_cur["ประเภทสินค้า"] = df_cur["ชื่อสินค้า"].map(st.session_state.cat_map)
            items_cur = df_cur[df_cur["ชื่อสินค้า"].notna()].copy()
            sum_cur   = make_summary(items_cur)

            df_prv = st.session_state.df_prev.copy()
            df_prv["ประเภทสินค้า"] = df_prv["ชื่อสินค้า"].map(st.session_state.cat_map_prev)
            items_prv = df_prv[df_prv["ชื่อสินค้า"].notna()].copy()
            sum_prv   = make_summary(items_prv)

            lbl_cur = items_cur["เดือน"].iloc[0] if len(items_cur) else "ปัจจุบัน"
            lbl_prv = items_prv["เดือน"].iloc[0] if len(items_prv) else "ก่อนหน้า"

            # metrics เปรียบเทียบ
            tot_cur = float(items_cur["ยอดรวมสินค้า"].sum())
            tot_prv = float(items_prv["ยอดรวมสินค้า"].sum())
            chg_pct = ((tot_cur-tot_prv)/tot_prv*100) if tot_prv>0 else 0
            arrow   = "📈" if chg_pct>=0 else "📉"
            color   = "#2E7D32" if chg_pct>=0 else "#C62828"

            m1,m2,m3 = st.columns(3)
            m1.markdown(f'<div class="mcard"><div class="icon">📅</div>'
                        f'<div class="num" style="font-size:1.1rem">{lbl_prv}</div>'
                        f'<div class="lbl">เดือนก่อน</div></div>', unsafe_allow_html=True)
            m2.markdown(f'<div class="mcard"><div class="icon">📅</div>'
                        f'<div class="num" style="font-size:1.1rem">{lbl_cur}</div>'
                        f'<div class="lbl">เดือนปัจจุบัน</div></div>', unsafe_allow_html=True)
            m3.markdown(f'<div class="mcard"><div class="icon">{arrow}</div>'
                        f'<div class="num" style="color:{color}">{chg_pct:+.1f}%</div>'
                        f'<div class="lbl">เปลี่ยนแปลงยอด</div></div>', unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            sec("เปรียบเทียบตามประเภทสินค้า","📊")

            # merge
            merged = sum_cur.rename(columns={"จำนวนรายการ":"จำนวน_ปัจจุบัน","ยอดรวม":"ยอด_ปัจจุบัน"}).merge(
                sum_prv.rename(columns={"จำนวนรายการ":"จำนวน_เดือนก่อน","ยอดรวม":"ยอด_เดือนก่อน"}),
                on="ประเภทสินค้า", how="outer").fillna(0)
            merged["เปลี่ยนแปลง%"] = merged.apply(
                lambda r: f'{((r["ยอด_ปัจจุบัน"]-r["ยอด_เดือนก่อน"])/r["ยอด_เดือนก่อน"]*100):+.1f}%'
                if r["ยอด_เดือนก่อน"]>0 else ("ใหม่" if r["ยอด_ปัจจุบัน"]>0 else "-"), axis=1)

            for _,row in merged.iterrows():
                cat  = row["ประเภทสินค้า"]
                meta = CATEGORIES.get(cat,{"icon":"📦","color":"#9E9E9E"})
                pct  = str(row["เปลี่ยนแปลง%"])
                up   = pct.startswith("+")
                ac   = "#2E7D32" if up else ("#C62828" if pct.startswith("-") else "#666")
                ab   = "E8F5E9" if up else ("FFEBEE" if pct.startswith("-") else "F5F5F5")
                st.markdown(
                    f'<div class="compare-card">'
                    f'<div style="display:flex;justify-content:space-between;align-items:center">'
                    f'<div style="display:flex;align-items:center;gap:8px">'
                    f'<span style="font-size:1.5rem">{meta["icon"]}</span>'
                    f'<div><div style="font-weight:700;color:#333">{cat}</div>'
                    f'<div style="font-size:.78rem;color:#888">'
                    f'{lbl_prv}: ฿{row["ยอด_เดือนก่อน"]:,.0f} → {lbl_cur}: ฿{row["ยอด_ปัจจุบัน"]:,.0f}</div>'
                    f'</div></div>'
                    f'<span style="background:#{ab};color:{ac};padding:4px 14px;border-radius:99px;'
                    f'font-weight:700;font-size:.9rem">{pct}</span>'
                    f'</div></div>', unsafe_allow_html=True)

            # Top branch เปรียบเทียบ
            st.markdown("<br>", unsafe_allow_html=True)
            sec("สาขาที่น่าสนใจ (ยอดเปลี่ยนแปลงมากสุด)","🏪")

            br_cur = items_cur.groupby("รหัสสาขา")["ยอดรวมสินค้า"].sum().reset_index()
            br_prv = items_prv.groupby("รหัสสาขา")["ยอดรวมสินค้า"].sum().reset_index()
            br_mrg = br_cur.merge(br_prv, on="รหัสสาขา", how="outer", suffixes=("_cur","_prv")).fillna(0)
            br_mrg["chg"] = br_mrg["ยอดรวมสินค้า_cur"] - br_mrg["ยอดรวมสินค้า_prv"]
            br_mrg["chg%"] = br_mrg.apply(
                lambda r: (r["chg"]/r["ยอดรวมสินค้า_prv"]*100) if r["ยอดรวมสินค้า_prv"]>0 else 100, axis=1)
            br_top = br_mrg.reindex(br_mrg["chg%"].abs().sort_values(ascending=False).index).head(5)

            for _,row in br_top.iterrows():
                bid  = int(row["รหัสสาขา"])
                up   = row["chg"]>=0
                icon = "📈" if up else "📉"
                color2 = "#2E7D32" if up else "#C62828"
                bg2    = "#E8F5E9" if up else "#FFEBEE"
                top_cat_cur = ""
                tc = items_cur[items_cur["รหัสสาขา"]==bid]
                if len(tc):
                    top_cat_cur = tc.groupby("ประเภทสินค้า")["ชื่อสินค้า"].count().idxmax()
                meta2 = CATEGORIES.get(top_cat_cur,{"icon":"📦"})
                st.markdown(
                    f'<div class="compare-card" style="border-left-color:{color2}">'
                    f'<div style="display:flex;justify-content:space-between;align-items:center">'
                    f'<div><div style="font-weight:700;font-size:1rem">🏪 สาขา {bid}</div>'
                    f'<div style="font-size:.78rem;color:#888">ประเภทนิยม: {meta2["icon"]} {top_cat_cur}</div>'
                    f'<div style="font-size:.78rem;color:#888">'
                    f'{lbl_prv}: ฿{row["ยอดรวมสินค้า_prv"]:,.0f} → {lbl_cur}: ฿{row["ยอดรวมสินค้า_cur"]:,.0f}</div>'
                    f'</div>'
                    f'<div style="text-align:right">'
                    f'<div style="font-size:1.8rem">{icon}</div>'
                    f'<span style="background:{bg2};color:{color2};padding:3px 12px;border-radius:99px;'
                    f'font-weight:700;font-size:.88rem">{row["chg%"]:+.1f}%</span>'
                    f'</div></div></div>', unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            st.info("💡 **วิเคราะห์เพิ่มเติม:** ดาวน์โหลด Excel ในแท็บแรก จะมี Sheet เปรียบเทียบพร้อมกราฟ "
                    "และตาราง Top 10 สาขาแยกตามเดือน")
        else:
            st.markdown("<br>", unsafe_allow_html=True)
            st.info("📌 อัปโหลดและจำแนกไฟล์ทั้งสองเดือนก่อนเพื่อดูผลเปรียบเทียบ")

st.markdown('<div class="footer">🛒 CJ Smart Scan · Powered by Gemini AI · Made with ❤️</div>',
            unsafe_allow_html=True)
