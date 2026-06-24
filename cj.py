import streamlit as st
import pandas as pd
import json, re, io, urllib.request
import numpy as np
from openpyxl import load_workbook, Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabel

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
    "Bao Cafe":         {"icon":"☕","color":"#C2185B","hex":"C2185B"},
    "อาหารพร้อมทาน":   {"icon":"🍱","color":"#F57C00","hex":"F57C00"},
    "เครื่องดื่ม":      {"icon":"🥤","color":"#1976D2","hex":"1976D2"},
    "ขนมขบเคี้ยว":     {"icon":"🍿","color":"#7B1FA2","hex":"7B1FA2"},
    "ของใช้ส่วนตัว":   {"icon":"🧴","color":"#00897B","hex":"00897B"},
    "ของใช้ในบ้าน":    {"icon":"🏠","color":"#558B2F","hex":"558B2F"},
    "สินค้าเบ็ดเตล็ด": {"icon":"🔋","color":"#6D4C41","hex":"6D4C41"},
    "บริการและอื่นๆ":  {"icon":"📱","color":"#455A64","hex":"455A64"},
}

RANK_MEDAL = ["🥇","🥈","🥉","4","5","6","7","8","9","10"]
RANK_BG    = ["FFD700","C0C0C0","CD7F32","FF6B6B","FF9F43",
              "48DBFB","1DD1A1","A29BFE","FD79A8","636E72"]

SYSTEM_PROMPT = """คุณคือระบบจำแนกประเภทสินค้าร้านสะดวกซื้อ CJ Express
วิเคราะห์จากชื่อสินค้าเท่านั้น
หมวดหมู่ที่ใช้ได้:
1. Bao Cafe → เครื่องดื่มร้าน Bao (ขึ้นต้น Bao_ หรือ Bac_)
2. อาหารพร้อมทาน → ปัง แซนด์วิช ไส้กรอก พุดดิ้ง ข้าวเกรียบ ไข่ แฮม
3. เครื่องดื่ม → น้ำดื่ม นม ชา กาแฟทั่วไป เบียร์ สุรา น้ำหวาน คอลลาเจน
4. ขนมขบเคี้ยว → มันฝรั่ง ลูกอม ช็อกโกแลต ไอศกรีม ถั่ว ขนมกรอบ
5. ของใช้ส่วนตัว → ยาสีฟัน แชมพู สบู่ ครีม โฟมล้างหน้า เครื่องสำอาง ยา
6. ของใช้ในบ้าน → ทิชชู ผงซักฟอก น้ำยาล้างจาน ถุงขยะ น้ำยาปรับผ้า
7. สินค้าเบ็ดเตล็ด → ถ่านไฟฉาย ไฟแช็ก เครื่องเขียน เทปลบ
8. บริการและอื่นๆ → เติมเงิน จ่ายบิล ส่วนลด แก้วลูกค้า
ตอบ JSON เท่านั้น: {"ชื่อสินค้า":"หมวดหมู่"}"""

# ─── HELPERS ───
def get_api_key():
    import os
    try:
        k = st.secrets["GEMINI_API_KEY"]
        if k and str(k).strip(): return str(k).strip()
    except: pass
    return os.environ.get("GEMINI_API_KEY","").strip()

def claude_classify(products, api_key):
    """เรียก Gemini API วิเคราะห์ประเภทสินค้า"""
    product_list = "\n".join(f"- {p}" for p in products)
    prompt = SYSTEM_PROMPT + "\n\nวิเคราะห์ประเภทสินค้าต่อไปนี้:\n" + product_list

    payload = json.dumps({
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.1,
            "maxOutputTokens": 2000,
        },
    }, ensure_ascii=False).encode("utf-8")

    url = (f"https://generativelanguage.googleapis.com/v1beta/models/"
           f"gemini-2.0-flash:generateContent?key={api_key}")
    req = urllib.request.Request(
        url, data=payload,
        headers={"Content-Type": "application/json; charset=utf-8"},
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            data = json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8")
        raise Exception(f"API {e.code}: {err_body[:300]}")

    raw = data["candidates"][0]["content"]["parts"][0]["text"]
    raw = re.sub(r"```json|```", "", raw).strip()
    return json.loads(raw)

# ─── RULE-BASED FALLBACK (ไม่ต้องใช้ API) ───
RULE_KEYWORDS = {
    "Bao Cafe":        ["bao_","bao ","bac_","bac ","เอสเปรสโซ","ลาเต้","อเมริกาโน","มัทฉะ","อาราบีก้า"],
    "อาหารพร้อมทาน":  ["ปัง","แซนด์วิช","ไส้กรอก","พุดดิ้ง","ข้าวเกรียบ","ผงชูรส","ไข่","แฮม","มินิบัน","เบเกอรี"],
    "เครื่องดื่ม":     ["น้ํา","น้ำ","นม","เบียร์","สุรา","โคล่า","โซดา","คอลลาเจน","โพรไบโอ","ชูกำลัง","เฮล","วิตซี","ซีวิท"],
    "ขนมขบเคี้ยว":    ["มันฝรั่ง","ลูกอม","ช็อก","ไอศกรีม","ถั่ว","เมล็ด","อัลมอนด์","ข้าวโพด","ทวิสโก้","คอนเน่","บันบัน","ซาซ่า","กินดะ","โรซี่","โคอาล่า","ตะวันขนม","ฟู่อารี่","เทสโต","เลยร้อย","เลยร็อค"],
    "ของใช้ส่วนตัว":  ["ยาสีฟัน","แชมพู","สบู่","ครีม","ผ้าอนามัย","โฟม","รองพื้น","แปรง","เดนทีน","บีโอเร","การ์นิเย่","สกินแสบ","เจเล","แพนดิน","สุภาภรณ์","เอล สมูธ","mbl","bigsml"],
    "ของใช้ในบ้าน":   ["ทิชชู","ผงซัก","น้ำยา","ถุงขยะ","ดาวน์นี่","วิกซอล","ปรับผ้า"],
    "สินค้าเบ็ดเตล็ด":["ถ่าน","ไฟแช็ก","ยากันยุง","เครื่องเขียน","เทปลบ","พานาโซนิค","ดราช่าง","เรนเจอร์"],
    "บริการและอื่นๆ": ["เติมเงิน","จ่ายบิล","ซิมการ์ด","บัตรเติม","ส่วนลด","แก้วลูกค้า","00"],
}

def rule_classify(name: str) -> str:
    n = str(name).lower()
    for cat, kws in RULE_KEYWORDS.items():
        if any(k.lower() in n for k in kws):
            return cat
    return ""   # ไม่รู้ → ส่งให้ AI

def classify_all(products, api_key):
    import time, hashlib

    # ── ขั้น 1: rule-based ก่อน ──
    result   = {}
    need_ai  = []
    for p in products:
        cat = rule_classify(p)
        if cat:
            result[p] = cat
        else:
            need_ai.append(p)

    total = len(products)
    bar    = st.progress(0)
    status = st.empty()
    status.markdown(f"⚡ Rule-based: จำแนกได้ **{len(result)}/{total}** รายการ "
                    f"| ส่ง AI อีก **{len(need_ai)}** รายการ")
    bar.progress(len(result)/total if total else 1)

    if not need_ai:
        bar.empty(); status.empty()
        return result

    # ── ขั้น 2: ส่ง AI เฉพาะที่ rule ไม่รู้ ──
    batch_size = 30
    batches    = [need_ai[i:i+batch_size] for i in range(0, len(need_ai), batch_size)]
    done_ai    = 0

    for i, batch in enumerate(batches):
        status.markdown(f"🤖 AI วิเคราะห์ **{done_ai}/{len(need_ai)}** รายการ "
                        f"(batch {i+1}/{len(batches)})...")
        retry = 0
        while retry < 3:
            try:
                result.update(claude_classify(batch, api_key))
                done_ai += len(batch)
                break
            except Exception as e:
                err = str(e)
                if "429" in err or "quota" in err.lower():
                    wait = (retry + 1) * 20
                    status.markdown(f"⏳ Rate limit — รอ **{wait} วิ** (retry {retry+1}/3)...")
                    time.sleep(wait)
                    retry += 1
                else:
                    # error อื่น → fallback rule แล้วไปต่อ
                    st.warning(f"Batch {i+1}: {err[:150]}")
                    for p in batch: result[p] = rule_classify(p) or "สินค้าเบ็ดเตล็ด"
                    done_ai += len(batch)
                    break
        else:
            # retry หมด → rule fallback
            for p in batch: result[p] = rule_classify(p) or "สินค้าเบ็ดเตล็ด"
            done_ai += len(batch)

        bar.progress((len(result))/total)
        if i < len(batches)-1:
            time.sleep(2)

    bar.empty(); status.empty()
    return result

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
    df["วันที่_dt"]   = pd.to_datetime(df["วันที่"], format="%d/%m/%Y", errors="coerce")
    df["เดือน_sort"]  = df["วันที่_dt"].dt.to_period("M").astype(str)
    df["เดือน_label"] = df["วันที่_dt"].dt.strftime("%b %Y")
    # top category per branch per month
    cat_grp = df.groupby(["เดือน_sort","รหัสสาขา","ประเภทสินค้า"])["ชื่อสินค้า"].count().reset_index()
    cat_grp = cat_grp.sort_values("ชื่อสินค้า",ascending=False).drop_duplicates(["เดือน_sort","รหัสสาขา"])
    cat_grp = cat_grp.rename(columns={"ชื่อสินค้า":"_c","ประเภทสินค้า":"top_cat"})
    grp = df.groupby(["เดือน_sort","เดือน_label","รหัสสาขา"])
    cnt = grp["ชื่อสินค้า"].count(); tot = grp["ยอดรวมสินค้า"].sum()
    m = pd.concat([cnt,tot],axis=1).reset_index()
    m.columns = ["เดือน_sort","เดือน_label","รหัสสาขา","จำนวนรายการ","ยอดรวม"]
    m = m.merge(cat_grp[["เดือน_sort","รหัสสาขา","top_cat"]], on=["เดือน_sort","รหัสสาขา"], how="left")
    return m.sort_values(["เดือน_sort","ยอดรวม"], ascending=[True,False])

# ─── EXCEL STYLES ───
def thin_border():
    s = Side(style="thin", color="DDDDDD")
    return Border(left=s, right=s, top=s, bottom=s)

def write_header(ws, row, ncols, title, bg="C9184A"):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=title)
    c.font      = Font(bold=True, color="FFFFFF", size=13)
    c.fill      = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 28

def write_col_headers(ws, row, headers):
    pinks = ["F8BBD0","FCE4EC","FFE0B2","E8F5E9","E3F2FD","EDE7F6","FFF8E1","F3E5F5"]
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=j, value=h)
        c.font      = Font(bold=True, color="555555", size=11)
        c.fill      = PatternFill("solid", fgColor=pinks[j % len(pinks)])
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = thin_border()
    ws.row_dimensions[row].height = 22

def stripe_row(ws, row, ncols, even=True):
    bg = "FFF5F5" if even else "FFFFFF"
    for col in range(1, ncols+1):
        c = ws.cell(row=row, column=col)
        c.fill      = PatternFill("solid", fgColor=bg)
        c.border    = thin_border()
        c.alignment = Alignment(vertical="center")

# ─── EXCEL BUILDER ───
def build_excel(df, summary, branch_df, map_df, items):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer,        sheet_name="ข้อมูลทั้งหมด",  index=False, startrow=2)
        summary.to_excel(writer,   sheet_name="สรุปตามประเภท",  index=False, startrow=2)
        branch_df.to_excel(writer, sheet_name="สรุปตามสาขา",    index=False, startrow=2)
        map_df.to_excel(writer,    sheet_name="mapping สินค้า", index=False, startrow=2)
        writer.book.create_sheet("กราฟ & ตารางสรุป")
    buf.seek(0)
    wb = load_workbook(buf)

    cat_hex = {c: CATEGORIES.get(c,{"hex":"9E9E9E"})["hex"] for c in summary["ประเภทสินค้า"]}

    # ── Sheet 1 ──
    ws1 = wb["ข้อมูลทั้งหมด"]
    write_header(ws1, 1, len(df.columns), "📋 ข้อมูลสินค้าทั้งหมด — วิเคราะห์ประเภทด้วย AI")
    write_col_headers(ws1, 3, df.columns.tolist())
    cat_col_idx = df.columns.tolist().index("ประเภทสินค้า") + 1
    for r in range(4, ws1.max_row+1):
        stripe_row(ws1, r, len(df.columns), r%2==0)
        cv = ws1.cell(row=r, column=cat_col_idx).value
        if cv and cv in cat_hex:
            h = cat_hex[cv]
            ws1.cell(row=r, column=cat_col_idx).fill = PatternFill("solid", fgColor=h+"33")
            ws1.cell(row=r, column=cat_col_idx).font = Font(bold=True, color=h)
    for col in ws1.columns:
        ws1.column_dimensions[get_column_letter(col[0].column)].width = 18
    ws1.freeze_panes = "A4"

    # ── Sheet 2 ──
    ws2 = wb["สรุปตามประเภท"]
    write_header(ws2, 1, len(summary.columns), "📊 สรุปจำนวนสินค้าตามประเภท")
    write_col_headers(ws2, 3, summary.columns.tolist())
    for r in range(4, ws2.max_row+1):
        stripe_row(ws2, r, len(summary.columns), r%2==0)
        cv = ws2.cell(row=r, column=1).value
        if cv and cv in cat_hex:
            h = cat_hex[cv]
            for c in range(1, len(summary.columns)+1):
                ws2.cell(row=r, column=c).fill = PatternFill("solid", fgColor=h+"22")
            ws2.cell(row=r, column=1).font = Font(bold=True, color=h)
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
    write_header(ws4, 1, 2, "🏷️ ตารางจับคู่สินค้า → ประเภท")
    write_col_headers(ws4, 3, ["ชื่อสินค้า","ประเภทสินค้า"])
    for r in range(4, ws4.max_row+1):
        stripe_row(ws4, r, 2, r%2==0)
        cv = ws4.cell(row=r, column=2).value
        if cv and cv in cat_hex:
            h = cat_hex[cv]
            ws4.cell(row=r, column=2).fill = PatternFill("solid", fgColor=h+"33")
            ws4.cell(row=r, column=2).font = Font(bold=True, color=h)
    ws4.column_dimensions["A"].width = 35
    ws4.column_dimensions["B"].width = 22

    # ── Sheet 5: กราฟ & ตาราง ──
    ws5 = wb["กราฟ & ตารางสรุป"]
    ws5.sheet_view.showGridLines = False
    ws5.sheet_properties.tabColor = "C9184A"
    for col_letter in [get_column_letter(i) for i in range(1,20)]:
        ws5.column_dimensions[col_letter].width = 14

    # Title
    write_header(ws5, 1, 16, "📈 กราฟ & ตารางสรุป — CJ Express Smart Scan")
    ws5.row_dimensions[1].height = 32

    # === BAR CHART (จำนวนรายการตามประเภท) ===
    # เขียนข้อมูล summary ลง hidden area (col P-Q)
    data_start_row = 3
    ws5.cell(row=data_start_row, column=17, value="ประเภทสินค้า")
    ws5.cell(row=data_start_row, column=18, value="จำนวน")
    for i, (_, row) in enumerate(summary.iterrows(), 1):
        ws5.cell(row=data_start_row+i, column=17, value=row["ประเภทสินค้า"])
        ws5.cell(row=data_start_row+i, column=18, value=int(row["จำนวนรายการ"]))

    n_cats = len(summary)
    bar = BarChart()
    bar.type    = "bar"
    bar.grouping= "clustered"
    bar.title   = "จำนวนรายการสินค้าตามประเภท"
    bar.style   = 10
    bar.y_axis.title = "จำนวนรายการ"
    bar.x_axis.title = "ประเภทสินค้า"
    bar.width   = 18; bar.height = 12

    data_ref = Reference(ws5, min_col=18, min_row=data_start_row,
                          max_row=data_start_row+n_cats)
    cats_ref = Reference(ws5, min_col=17, min_row=data_start_row+1,
                          max_row=data_start_row+n_cats)
    bar.add_data(data_ref, titles_from_data=True)
    bar.set_categories(cats_ref)
    bar.series[0].graphicalProperties.solidFill   = "FF4D6D"
    bar.series[0].graphicalProperties.line.solidFill = "C9184A"
    ws5.add_chart(bar, "A3")

    # === PIE CHART (สัดส่วน) ===
    pie = PieChart()
    pie.title  = "สัดส่วนประเภทสินค้า"
    pie.style  = 10
    pie.width  = 14; pie.height = 12
    pie_data = Reference(ws5, min_col=18, min_row=data_start_row,
                          max_row=data_start_row+n_cats)
    pie_cats = Reference(ws5, min_col=17, min_row=data_start_row+1,
                          max_row=data_start_row+n_cats)
    pie.add_data(pie_data, titles_from_data=True)
    pie.set_categories(pie_cats)
    pie.dataLabels = DataLabel(showPercent=True, showVal=False, showCatName=False)
    ws5.add_chart(pie, "L3")

    # === ตาราง TOP 10 สาขา รายเดือน ===
    monthly = make_monthly_top(items)
    months  = monthly["เดือน_sort"].unique()

    tbl_start_row = 24   # เริ่มหลังกราฟ
    current_row   = tbl_start_row

    # Header ใหญ่
    ws5.merge_cells(start_row=current_row, start_column=1,
                    end_row=current_row, end_column=8)
    hc = ws5.cell(row=current_row, column=1,
                  value="🏆 Top 10 สาขาขายดี แยกตามเดือน")
    hc.font      = Font(bold=True, color="FFFFFF", size=14)
    hc.fill      = PatternFill("solid", fgColor="880E4F")
    hc.alignment = Alignment(horizontal="center", vertical="center")
    ws5.row_dimensions[current_row].height = 30
    current_row += 1

    col_headers = ["อันดับ","รหัสสาขา","เดือน","ยอดรวม (บาท)","จำนวนรายการ","ประเภทยอดนิยม"]
    col_widths_tbl = [10, 14, 14, 18, 18, 22]
    for j, (h, w) in enumerate(zip(col_headers, col_widths_tbl), 1):
        ws5.column_dimensions[get_column_letter(j)].width = w

    for month_key in months:
        mdata = monthly[monthly["เดือน_sort"]==month_key].head(10)
        month_label = mdata["เดือน_label"].iloc[0]

        # Month subheader
        ws5.merge_cells(start_row=current_row, start_column=1,
                        end_row=current_row, end_column=6)
        mc = ws5.cell(row=current_row, column=1, value=f"📅  {month_label}")
        mc.font      = Font(bold=True, color="FFFFFF", size=12)
        mc.fill      = PatternFill("solid", fgColor="C9184A")
        mc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws5.row_dimensions[current_row].height = 24
        current_row += 1

        # Column headers
        header_fills = ["FFE0B2","FFF3E0","E3F2FD","E8F5E9","FCE4EC","EDE7F6"]
        for j, (h, hfill) in enumerate(zip(col_headers, header_fills), 1):
            c = ws5.cell(row=current_row, column=j, value=h)
            c.font      = Font(bold=True, color="333333", size=10)
            c.fill      = PatternFill("solid", fgColor=hfill)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border    = thin_border()
        ws5.row_dimensions[current_row].height = 20
        current_row += 1

        # Data rows
        for i, (_, row) in enumerate(mdata.iterrows()):
            rank_hex = RANK_BG[i] if i < len(RANK_BG) else "AAAAAA"
            medal    = RANK_MEDAL[i] if i < len(RANK_MEDAL) else str(i+1)
            cat      = str(row.get("top_cat",""))
            cat_hex_val = CATEGORIES.get(cat,{"hex":"9E9E9E"})["hex"]
            bg_row   = "FFF0F3" if i%2==0 else "FFFFFF"

            vals = [medal, str(int(row["รหัสสาขา"])), str(row["เดือน_label"]),
                    f'{row["ยอดรวม"]:,.2f}', str(int(row["จำนวนรายการ"])), cat]

            for j, val in enumerate(vals, 1):
                c = ws5.cell(row=current_row, column=j, value=val)
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border    = thin_border()
                if j == 1:  # rank
                    c.fill = PatternFill("solid", fgColor=rank_hex)
                    c.font = Font(bold=True, color="FFFFFF", size=11)
                elif j == 6:  # category
                    c.fill = PatternFill("solid", fgColor=cat_hex_val+"44")
                    c.font = Font(bold=True, color=cat_hex_val, size=10)
                else:
                    c.fill = PatternFill("solid", fgColor=bg_row)
                    c.font = Font(color="333333", size=10)
            ws5.row_dimensions[current_row].height = 20
            current_row += 1

        current_row += 1  # spacer between months

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

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

api_key = get_api_key()
if not api_key:
    st.error("⚠️ ไม่พบ **GEMINI_API_KEY**")
    st.info("ไปที่ App Settings → Secrets แล้วใส่:\n```\nGEMINI_API_KEY = \"AIza...\"\n```")
    with st.expander("🔍 Debug"):
        try: st.write("Keys:", list(st.secrets.keys()))
        except Exception as e: st.write("อ่าน secrets ไม่ได้:", e)
    st.stop()

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
            st.session_state.df       = df_raw
            st.session_state.analyzed = False
            st.session_state.cat_map  = {}
            n_items  = int(df_raw["ชื่อสินค้า"].notna().sum())
            n_branch = int(df_raw["รหัสสาขา"].dropna().nunique())
            st.success(f"✅ โหลดสำเร็จ! พบ **{n_items}** รายการ จาก **{n_branch}** สาขา")
        except Exception as e:
            st.error(f"❌ {e}"); st.stop()

    if st.session_state.df is not None and not st.session_state.analyzed:
        if st.button("🤖  วิเคราะห์ประเภทสินค้า", type="primary", use_container_width=True):
            products = st.session_state.df["ชื่อสินค้า"].dropna().unique().tolist()
            st.session_state.cat_map  = classify_all(products, api_key)
            st.session_state.analyzed = True
            st.rerun()

    if st.session_state.analyzed:
        sec("หมวดหมู่สินค้า", "🏷️")
        for name, meta in CATEGORIES.items():
            st.markdown(
                f'<div style="display:flex;align-items:center;gap:10px;padding:5px 0">'
                f'<span style="font-size:1.3rem">{meta["icon"]}</span>'
                f'<span style="background:{meta["color"]};color:white;padding:3px 12px;'
                f'border-radius:99px;font-size:.8rem;font-weight:600">{name}</span>'
                f'</div>', unsafe_allow_html=True)

with right:
    if st.session_state.analyzed and st.session_state.df is not None:
        df = st.session_state.df.copy()
        df["ประเภทสินค้า"] = df["ชื่อสินค้า"].map(st.session_state.cat_map)
        items = df[df["ชื่อสินค้า"].notna()].copy()

        summary_df = make_summary(items)
        branch_df  = make_branch_summary(items)
        map_df     = pd.DataFrame(sorted(st.session_state.cat_map.items(), key=lambda x:x[1]),
                                  columns=["ชื่อสินค้า","ประเภทสินค้า"])

        excel_bytes = build_excel(df, summary_df, branch_df, map_df, items)
        st.download_button(
            "⬇️  ดาวน์โหลด Excel (กราฟ + ตารางสรุป Top 10 สาขา)",
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
            col.markdown(f'<div class="mcard"><div class="icon">{icon}</div>'
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
            pct  = cnt/max_cnt*100
            st.markdown(
                f'<div class="cat-card">'
                f'<div class="cat-icon">{meta["icon"]}</div>'
                f'<div style="flex:1"><div class="cat-name">{cat}</div>'
                f'<div class="cat-count">฿{tot:,.0f}</div>'
                f'<div class="bar-wrap" style="margin-top:6px">'
                f'<div class="bar-fill" style="width:{pct}%;background:{meta["color"]}"></div>'
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
        search = st.text_input("🔍 ค้นหา", placeholder="พิมพ์ชื่อสินค้า...")

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
                st.dataframe(bd[["ประเภทสินค้า","จำนวนรายการ","ยอดรวม"]]
                             .style.format({"ยอดรวม":"฿{:,.2f}"}),
                             use_container_width=True, hide_index=True, height=260)

st.markdown('<div class="footer">🛒 CJ Smart Scan · Powered by Claude AI · Made with ❤️</div>',
            unsafe_allow_html=True)
