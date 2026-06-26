import streamlit as st
import pandas as pd
import json, re, io, urllib.request
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference

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
    position: relative; overflow: hidden; box-shadow: 0 8px 32px rgba(201,24,74,.25);
}
.hero-title { font-family:'Fredoka One','Sarabun',sans-serif !important; font-size:2.4rem; font-weight:900; color:white; margin:0; }
.hero-sub { color:rgba(255,255,255,.88); font-size:1rem; margin:8px 0 0; }
.hero-mascot { position:absolute; right:40px; top:50%; transform:translateY(-50%); font-size:6rem; }
.sec-title { display:flex; align-items:center; gap:10px; font-size:1.1rem; font-weight:700; color:#C9184A; margin:22px 0 12px; }
.sec-title .dot { width:5px; height:22px; background:#FF4D6D; border-radius:3px; }
.mcard { background:white; border-radius:16px; padding:20px 14px; text-align:center; box-shadow:0 3px 12px rgba(0,0,0,.07); }
.mcard .icon { font-size:2.2rem; line-height:1; margin-bottom:4px; }
.mcard .num  { font-size:1.8rem; font-weight:800; color:#C9184A; }
.mcard .lbl  { font-size:.78rem; color:#999; margin-top:4px; font-weight:500; }
.cat-card { background:#FFFBFC; border-radius:14px; padding:13px 18px; margin-bottom:8px;
    display:flex; align-items:center; gap:12px; box-shadow:0 1px 6px rgba(0,0,0,.05); border:1px solid #FFF0F3; }
.cat-card .cat-icon { font-size:1.8rem; }
.cat-card .cat-name { font-weight:700; font-size:.9rem; color:#333; }
.cat-card .cat-count { font-size:.78rem; color:#aaa; }
.cat-card .bar-wrap { flex:1; background:#f0f0f0; border-radius:99px; height:7px; overflow:hidden; }
.cat-card .bar-fill { height:100%; border-radius:99px; }
.cat-card .cat-num { font-weight:800; font-size:1rem; color:#C9184A; min-width:32px; text-align:right; }
.compare-card { background:white; border-radius:16px; padding:18px 20px; margin-bottom:10px;
    box-shadow:0 2px 10px rgba(0,0,0,.06); border-left:4px solid #FF4D6D; }
.stDownloadButton > button { background:linear-gradient(135deg,#FF6B6B,#C9184A) !important;
    color:white !important; border:none !important; border-radius:12px !important;
    padding:12px 24px !important; font-size:.95rem !important; font-weight:700 !important; }
.stButton > button[kind="primary"] { background:linear-gradient(135deg,#FF6B6B,#C9184A) !important;
    color:white !important; border:none !important; border-radius:12px !important; }
.footer { text-align:center; color:#ccc; font-size:.78rem; padding:28px 0 12px; }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════
# CONSTANTS — หมวดหมู่ตาม Product_Group_by_PMA
# ═══════════════════════════════════════════
CATEGORIES = {
    "Bao Cafe":          {"icon":"☕","color":"#C2185B","hex":"C2185B"},
    "Fresh Food":        {"icon":"🥗","color":"#2E7D32","hex":"2E7D32"},
    "Non Food":          {"icon":"🧴","color":"#1565C0","hex":"1565C0"},
    "Packaged Beverage": {"icon":"🥤","color":"#6A1B9A","hex":"6A1B9A"},
    "Processed Food":    {"icon":"🍿","color":"#E65100","hex":"E65100"},
    "Special Business":  {"icon":"⭐","color":"#AD1457","hex":"AD1457"},
    "ส่วนลด/โปรโมชั่น":  {"icon":"🏷️","color":"#00695C","hex":"00695C"},
}

RULE_KEYWORDS = {
    "Bao Cafe": [
        "bao_","bao ","bac_","bac ","เอสเปรสโซ","ลาเต้","อเมริกาโน","มัทฉะ",
        "ชาเขียวนมสด","โกโก้","คาปูชิโน","bao","8ao","b4o",
    ],
    "ส่วนลด/โปรโมชั่น": [
        "ส่วนลด","โปรโมชั่น","โปรโม","discount","promotion","แถม",
    ],
    "Fresh Food": [
        "ขนมปัง","แซนด์วิช","แซนด์","burger","hotdog","ไส้กรอก","ไก่ทอด","หมูทอด","เบเกอรี",
        "bakery","bread","นมพาสเจอ","pasteurized","ข้าวกล่อง","ข้าว","สลัด","salad",
        "meal box","พุดดิ้ง","ลูกชิ้น","เกาเหลา","ต้มยำ","แกง","โจ๊ก","grilled","sausage",
        "frozen","chilled","rtc","ready to cook","retort","warmed","toasted","cpg",
    ],
    "Non Food": [
        "บุหรี่","cigarette","สมุด","หนังสือ","book","magazine","หนังสือพิมพ์","newspaper",
        "ยาสีฟัน","แชมพู","สบู่","ครีม","โฟม","ผ้าอนามัย","แปรง","รองพื้น","เครื่องสำอาง",
        "personal care","houseware","ทิชชู","ผงซักฟอก","น้ำยา","ถุงขยะ","household",
        "sanitary","stationery","เครื่องเขียน","ปากกา","ดินสอ","เทปลบ","กรรไกร",
        "electronic","it ","โทรศัพท์","ชาร์จ","สายชาร์จ","หูฟัง","herbal",
        "ไฟแช็ก","lighter","ถ่าน","battery","ยากันยุง",
    ],
    "Packaged Beverage": [
        "เบียร์","beer","สุรา","liquor","วิสกี้","whisky","vodka","ไวน์","wine",
        "นมuht","uht","นมกล่อง","น้ำผลไม้","น้ำส้ม","น้ำองุ่น","น้ำมะขาม",
        "โค้ก","โคก","coke","cola","เป๊ปซี่","pepsi","สไปรท์","sprite","แฟนต้า",
        "โซดา","soda","น้ำอัดลม","csd","carbonated",
        "ไอศกรีม","ice cream","cornetto","wall","magnum","นํ้าแข็ง","น้ำแข็ง","ice",
        "คาราบาว","กระทิง","ชาร์จ","shark","m150","เรดบูล","redbull","energy drink",
        "sport drink","gatorade","100plus","เกลือแร่","โพลาริส",
        "non-carbonated","non carbonated","ชาเขียว","ชาดำ","น้ำชา","ชาพร้อมดื่ม",
        "โออิชิ","ลิปตัน","ไวตามิลค์","vitamilk","dutch","ดัชมิลล์",
    ],
    "Processed Food": [
        "มาม่า","บะหมี่","mie sedap","instant","สำเร็จรูป",
        "ซอส","น้ำปลา","น้ำตาล","เกลือ","cooking","canned","กระป๋อง","ปลากระป๋อง",
        "ลูกอม","candy","ช็อกโกแลต","chocolate","gummy","เยลลี่","jelly",
        "มันฝรั่ง","potato","เลย์","lay","pringle","ข้าวโพด","popcorn","ป๊อปคอร์น",
        "ถั่ว","nuts","อัลมอนด์","almond","ขนมกรอบ","snack","ทวิสโก้","คอนเน่",
        "ข้าวอบ","ปังกรอบ","wafer","เวเฟอร์","คุกกี้","cookie","cracker",
        "confectionery","thai snack","dry fruit","ลำไยอบ","มะม่วง","ผลไม้อบ",
    ],
    "Special Business": [
        "ยา","medicine","drug","paracetamol","ibuprofen","วิตามิน","vitamin",
        "อาหารเสริม","supplement","health","wellness","แอลกอฮอล์","alcohol","หน้ากาก","mask",
        "เติมเงิน","จ่ายบิล","ซิมการ์ด","บัตรเติม","service","commission","gp",
        "bellinee","kudsan","synergy","fresh bakery","social welfare","supply","catalog",
        "vegetable","ผัก","home & living","hot served","ผ้าพันแผล","พลาสเตอร์",
    ],
}

RANK_BG    = ["FFD700","C0C0C0","CD7F32","FF6B6B","FF9F43","48DBFB","1DD1A1","A29BFE","FD79A8","636E72"]
RANK_MEDAL = ["🥇","🥈","🥉","4","5","6","7","8","9","10"]

SYSTEM_PROMPT = """คุณคือระบบจำแนกประเภทสินค้าร้าน CJ Express ตามกลุ่ม PMA
วิเคราะห์จากชื่อสินค้าเท่านั้น ตอบ JSON เท่านั้น: {"ชื่อสินค้า":"หมวดหมู่"}

หมวดหมู่ที่ใช้ได้ (เลือกได้เฉพาะ 7 หมวดนี้):
1. Bao Cafe — เครื่องดื่ม Bao ทุกชนิด (ลาเต้ อเมริกาโน่ โกโก้ ฯลฯ)
2. Fresh Food — อาหารสด อาหารพร้อมทาน ขนมปัง แซนด์วิช ไส้กรอก นมพาสเจอร์ไรส์
3. Non Food — ของใช้ส่วนตัว ของใช้ในบ้าน เครื่องเขียน บุหรี่ อิเล็กทรอนิกส์
4. Packaged Beverage — เครื่องดื่มบรรจุขวด/กล่อง น้ำอัดลม นม UHT เบียร์ ไวน์ สุรา ไอศกรีม น้ำแข็ง เครื่องดื่มชูกำลัง
5. Processed Food — อาหารแปรรูป บะหมี่กึ่งสำเร็จรูป ขนมขบเคี้ยว ลูกอม ช็อกโกแลต ของหวาน
6. Special Business — ยา อาหารเสริม บริการ เติมเงิน ผัก เบเกอรี่พิเศษ
7. ส่วนลด/โปรโมชั่น — ส่วนลดและโปรโมชั่นเท่านั้น"""

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
        if any(k.lower() in n for k in kws):
            return cat
    return ""

def classify_rule_only(products):
    return {p: rule_classify(p) or "Processed Food" for p in products}

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
    need_ai = [p for p in products if not rule_classify(p)]
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
                st.warning(f"⚠️ Rate limit — ใช้ rule-based แทน")
                break
            else:
                st.warning(f"Batch {i+1}: {err[:120]}")
        bar.progress((i+1)/len(batches))
    bar.empty(); status.empty()
    return result

def load_excel(file):
    df = pd.read_excel(file, sheet_name="ใบเสร็จ")
    drop = [c for c in df.columns if "หมวด" in c or "ประเภท" in c or "category" in c.lower()]
    if drop:
        df = df.drop(columns=drop)
    df["วันที่_dt"]  = pd.to_datetime(df["วันที่"], format="%d/%m/%Y", errors="coerce")
    df["เดือน"]      = df["วันที่_dt"].dt.strftime("%b %Y")
    df["เดือน_sort"] = df["วันที่_dt"].dt.to_period("M").astype(str)
    return df

def parse_receipt(df):
    """แตกเลขที่ใบเสร็จ → เครื่องคิดเงิน + เลขที่ใบเสร็จ (ยอดลูกค้า)"""
    def extract(r):
        m = re.search(r'(N\d+)-(\d+)', str(r))
        if m:
            return m.group(1), int(m.group(2))
        return None, None
    df = df.copy()
    parsed        = df["เลขที่ใบเสร็จ"].apply(extract)
    df["เครื่อง"] = parsed.apply(lambda x: x[0])
    df["ลำดับใบเสร็จ"] = parsed.apply(lambda x: x[1])
    return df

def get_customer_count(df):
    """
    ยอดลูกค้าสะสมของแต่ละสาขา = max(ลำดับใบเสร็จ) ต่อเครื่อง รวมทุกเครื่อง
    เพราะแต่ละเครื่องนับลูกค้าแยกกัน
    """
    df2 = parse_receipt(df)
    # max ลำดับต่อ (สาขา, เครื่อง) = จำนวนลูกค้าของเครื่องนั้น
    per_machine = df2.groupby(["รหัสสาขา","เครื่อง"])["ลำดับใบเสร็จ"].max().reset_index()
    per_machine.columns = ["รหัสสาขา","เครื่อง","ยอดลูกค้า"]
    # รวมทุกเครื่องต่อสาขา
    per_branch  = per_machine.groupby("รหัสสาขา")["ยอดลูกค้า"].sum().reset_index()
    per_branch.columns = ["รหัสสาขา","ยอดลูกค้าสะสม"]
    return per_branch, per_machine

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

# ═══════════════════════════════════════════
# EXCEL BUILDER
# ═══════════════════════════════════════════
def thin_border():
    s = Side(style="thin", color="E8E8E8")
    return Border(left=s, right=s, top=s, bottom=s)

def write_header(ws, row, ncols, title, bg="C9184A"):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=title)
    c.font      = Font(bold=True, color="FFFFFF", size=13)
    c.fill      = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 28

def write_col_headers(ws, row, headers, bg="FCE4EC"):
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=j, value=h)
        c.font      = Font(bold=True, color="555555", size=10)
        c.fill      = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = thin_border()
    ws.row_dimensions[row].height = 20

def stripe_row(ws, row, ncols, even=True):
    bg = "FFF8FA" if even else "FFFFFF"
    for col in range(1, ncols+1):
        c = ws.cell(row=row, column=col)
        if not c.fill or c.fill.fgColor.rgb in ("00000000","FFFFFFFF","FFF8FA","FFFFFF"):
            c.fill = PatternFill("solid", fgColor=bg)
        c.border    = thin_border()
        c.alignment = Alignment(vertical="center")

def build_excel(df, summary, branch_df, map_df, items, cust_branch, cust_machine,
                df_prev=None, summary_prev=None):
    buf = io.BytesIO()
    cat_hex = {c: CATEGORIES.get(c,{"hex":"9E9E9E"})["hex"] for c in summary["ประเภทสินค้า"]}

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="ข้อมูลทั้งหมด",    index=False, startrow=2)
        summary.to_excel(writer,   sheet_name="สรุปตามประเภท",  index=False, startrow=2)
        branch_df.to_excel(writer, sheet_name="สรุปตามสาขา",    index=False, startrow=2)
        map_df.to_excel(writer,    sheet_name="mapping สินค้า",  index=False, startrow=2)
        cust_branch.to_excel(writer, sheet_name="ยอดลูกค้าสาขา", index=False, startrow=2)
        cust_machine.to_excel(writer, sheet_name="ยอดลูกค้าเครื่อง", index=False, startrow=2)
        writer.book.create_sheet("กราฟ & Top สาขา")
        if df_prev is not None:
            writer.book.create_sheet("เปรียบเทียบเดือน")

    buf.seek(0)
    wb = load_workbook(buf)

    # ── Sheet 1: ข้อมูลทั้งหมด ──
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
    write_header(ws2, 1, len(summary.columns), "📊 สรุปตามประเภทสินค้า (PMA Groups)")
    write_col_headers(ws2, 3, summary.columns.tolist())
    for r in range(4, ws2.max_row+1):
        cv  = ws2.cell(row=r, column=1).value
        h   = cat_hex.get(cv, "F5F5F5")
        bg  = h+"15"
        for c in range(1, len(summary.columns)+1):
            cell = ws2.cell(row=r, column=c)
            cell.fill      = PatternFill("solid", fgColor=bg)
            cell.border    = thin_border()
            cell.alignment = Alignment(vertical="center", horizontal="center")
        ws2.cell(row=r,column=1).font = Font(bold=True, color=h if h!="F5F5F5" else "333333")
    for col in ws2.columns:
        ws2.column_dimensions[get_column_letter(col[0].column)].width = 22

    # ── Sheet 3: สรุปตามสาขา ──
    ws3 = wb["สรุปตามสาขา"]
    write_header(ws3, 1, len(branch_df.columns), "🏪 สรุปตามสาขา")
    write_col_headers(ws3, 3, branch_df.columns.tolist())
    for r in range(4, ws3.max_row+1):
        stripe_row(ws3, r, len(branch_df.columns), r%2==0)
    for col in ws3.columns:
        ws3.column_dimensions[get_column_letter(col[0].column)].width = 20

    # ── Sheet 4: mapping สินค้า ──
    ws4 = wb["mapping สินค้า"]
    write_header(ws4, 1, 2, "🏷️ mapping สินค้า → ประเภท (PMA)")
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

    # ── Sheet 5: ยอดลูกค้าสาขา ──
    ws5 = wb["ยอดลูกค้าสาขา"]
    write_header(ws5, 1, len(cust_branch.columns), "👥 ยอดลูกค้าสะสมต่อสาขา", "1565C0")
    write_col_headers(ws5, 3, cust_branch.columns.tolist(), "BBDEFB")
    for r in range(4, ws5.max_row+1):
        stripe_row(ws5, r, len(cust_branch.columns), r%2==0)
        # highlight ยอดลูกค้า
        cv = ws5.cell(row=r, column=2).value
        if cv:
            ws5.cell(row=r, column=2).font = Font(bold=True, color="1565C0", size=11)
            ws5.cell(row=r, column=2).alignment = Alignment(horizontal="center", vertical="center")
    ws5.column_dimensions["A"].width = 15
    ws5.column_dimensions["B"].width = 25
    ws5.freeze_panes = "A4"

    # ── Sheet 6: ยอดลูกค้าเครื่อง ──
    ws6 = wb["ยอดลูกค้าเครื่อง"]
    write_header(ws6, 1, len(cust_machine.columns), "🖨️ ยอดลูกค้าสะสมแยกตามเครื่องคิดเงิน", "6A1B9A")
    write_col_headers(ws6, 3, cust_machine.columns.tolist(), "EDE7F6")
    for r in range(4, ws6.max_row+1):
        stripe_row(ws6, r, len(cust_machine.columns), r%2==0)
        cv = ws6.cell(row=r, column=3).value
        if cv:
            ws6.cell(row=r, column=3).font = Font(bold=True, color="6A1B9A", size=11)
            ws6.cell(row=r, column=3).alignment = Alignment(horizontal="center", vertical="center")
    for i, w in enumerate([15, 15, 25], 1):
        ws6.column_dimensions[get_column_letter(i)].width = w
    ws6.freeze_panes = "A4"

    # ── Sheet 7: กราฟ & Top สาขา ──
    ws7 = wb["กราฟ & Top สาขา"]
    ws7.sheet_view.showGridLines = False
    ws7.sheet_properties.tabColor = "C9184A"
    for i in range(1, 18):
        ws7.column_dimensions[get_column_letter(i)].width = 14
    write_header(ws7, 1, 16, "📈 กราฟ & Top 10 สาขา — ยอดลูกค้าสะสม")

    # data for charts (col Q-R)
    ws7.cell(row=3, column=17, value="ประเภท")
    ws7.cell(row=3, column=18, value="จำนวน")
    for i, (_, row) in enumerate(summary.iterrows(), 1):
        ws7.cell(row=3+i, column=17, value=row["ประเภทสินค้า"])
        ws7.cell(row=3+i, column=18, value=int(row["จำนวนรายการ"]))
    n = len(summary)

    bar = BarChart(); bar.type = "bar"; bar.grouping = "clustered"
    bar.title = "จำนวนรายการตามประเภท PMA"; bar.style = 10; bar.width = 18; bar.height = 12
    bar.add_data(Reference(ws7, min_col=18, min_row=3, max_row=3+n), titles_from_data=True)
    bar.set_categories(Reference(ws7, min_col=17, min_row=4, max_row=3+n))
    bar.series[0].graphicalProperties.solidFill = "FF4D6D"
    bar.series[0].graphicalProperties.line.solidFill = "C9184A"
    ws7.add_chart(bar, "A3")

    pie = PieChart(); pie.title = "สัดส่วนประเภทสินค้า PMA"; pie.style = 10; pie.width = 14; pie.height = 12
    pie.add_data(Reference(ws7, min_col=18, min_row=3, max_row=3+n), titles_from_data=True)
    pie.set_categories(Reference(ws7, min_col=17, min_row=4, max_row=3+n))
    ws7.add_chart(pie, "L3")

    # Top 10 ยอดลูกค้าต่อสาขา
    cr = 24
    write_header(ws7, cr, 6, "🏆 Top 10 สาขา — ยอดลูกค้าสะสมสูงสุด", bg="880E4F")
    ws7.row_dimensions[cr].height = 28; cr += 1
    write_col_headers(ws7, cr, ["อันดับ","รหัสสาขา","ยอดลูกค้าสะสม","(ทุกเครื่องรวมกัน)","",""], "FFE0B2")
    ws7.row_dimensions[cr].height = 20; cr += 1

    top10 = cust_branch.sort_values("ยอดลูกค้าสะสม", ascending=False).head(10)
    for i, (_, row) in enumerate(top10.iterrows()):
        rh  = RANK_BG[i] if i < len(RANK_BG) else "AAAAAA"
        med = RANK_MEDAL[i] if i < len(RANK_MEDAL) else str(i+1)
        bg  = "FFF8FA" if i%2==0 else "FFFFFF"
        vals = [med, str(int(row["รหัสสาขา"])), int(row["ยอดลูกค้าสะสม"]), "transaction", "", ""]
        for j, val in enumerate(vals, 1):
            c = ws7.cell(row=cr, column=j, value=val)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border    = thin_border()
            if j == 1:
                c.fill = PatternFill("solid", fgColor=rh); c.font = Font(bold=True, color="FFFFFF", size=10)
            elif j == 3:
                c.fill = PatternFill("solid", fgColor="E3F2FD"); c.font = Font(bold=True, color="1565C0", size=11)
            else:
                c.fill = PatternFill("solid", fgColor=bg); c.font = Font(color="333333", size=9)
        ws7.row_dimensions[cr].height = 20; cr += 1

    # Monthly top
    monthly = make_monthly_top(items)
    cr += 1
    write_header(ws7, cr, 6, "📅 Top 10 สาขา รายเดือน", bg="C9184A")
    ws7.row_dimensions[cr].height = 28; cr += 1

    for month_key in sorted(monthly["เดือน_sort"].unique()):
        mdata      = monthly[monthly["เดือน_sort"] == month_key].head(10)
        month_lbl  = mdata["เดือน"].iloc[0]
        ws7.merge_cells(start_row=cr, start_column=1, end_row=cr, end_column=6)
        mc = ws7.cell(row=cr, column=1, value=f"📅  {month_lbl}")
        mc.font = Font(bold=True, color="FFFFFF", size=11)
        mc.fill = PatternFill("solid", fgColor="C9184A")
        mc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws7.row_dimensions[cr].height = 22; cr += 1

        write_col_headers(ws7, cr, ["อันดับ","รหัสสาขา","เดือน","ยอดรวม","จำนวน","ประเภทนิยม"], "FFE0B2")
        ws7.row_dimensions[cr].height = 18; cr += 1

        for i, (_, row) in enumerate(mdata.iterrows()):
            rh  = RANK_BG[i] if i < len(RANK_BG) else "AAAAAA"
            cat = str(row.get("top_cat",""))
            ch  = CATEGORIES.get(cat, {"hex":"9E9E9E"})["hex"]
            bg  = "FFF8FA" if i%2==0 else "FFFFFF"
            vals = [RANK_MEDAL[i] if i<len(RANK_MEDAL) else str(i+1),
                    str(int(row["รหัสสาขา"])), str(row["เดือน"]),
                    f'{row["ยอดรวม"]:,.2f}', str(int(row["จำนวนรายการ"])), cat]
            for j, val in enumerate(vals, 1):
                c = ws7.cell(row=cr, column=j, value=val)
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border    = thin_border()
                if j == 1:
                    c.fill = PatternFill("solid", fgColor=rh); c.font = Font(bold=True, color="FFFFFF", size=10)
                elif j == 6:
                    c.fill = PatternFill("solid", fgColor=ch+"22"); c.font = Font(bold=True, color=ch, size=9)
                else:
                    c.fill = PatternFill("solid", fgColor=bg); c.font = Font(color="333333", size=9)
            ws7.row_dimensions[cr].height = 18; cr += 1
        cr += 1

    # ── Sheet 8: เปรียบเทียบ ──
    if df_prev is not None and summary_prev is not None:
        ws8 = wb["เปรียบเทียบเดือน"]
        ws8.sheet_view.showGridLines = False
        ws8.sheet_properties.tabColor = "1976D2"
        for i in range(1, 14):
            ws8.column_dimensions[get_column_letter(i)].width = 18
        write_header(ws8, 1, 10, "📊 เปรียบเทียบยอดขายระหว่างสองเดือน", "1565C0")
        ws8.row_dimensions[1].height = 28

        merged = summary.rename(columns={"จำนวนรายการ":"จำนวน_ปัจจุบัน","ยอดรวม":"ยอด_ปัจจุบัน"}).merge(
            summary_prev.rename(columns={"จำนวนรายการ":"จำนวน_เดือนก่อน","ยอดรวม":"ยอด_เดือนก่อน"}),
            on="ประเภทสินค้า", how="outer").fillna(0)
        merged["เปลี่ยนแปลง%"] = merged.apply(
            lambda r: f'{((r["ยอด_ปัจจุบัน"]-r["ยอด_เดือนก่อน"])/r["ยอด_เดือนก่อน"]*100):+.1f}%'
            if r["ยอด_เดือนก่อน"] > 0 else ("ใหม่" if r["ยอด_ปัจจุบัน"] > 0 else "-"), axis=1)

        hdrs = ["ประเภทสินค้า","จำนวน_เดือนก่อน","ยอด_เดือนก่อน","จำนวน_ปัจจุบัน","ยอด_ปัจจุบัน","เปลี่ยนแปลง%"]
        write_col_headers(ws8, 3, hdrs, "BBDEFB")
        for r_idx, (_, row) in enumerate(merged[hdrs].iterrows(), 4):
            cat = row["ประเภทสินค้า"]
            h   = CATEGORIES.get(cat, {"hex":"9E9E9E"})["hex"]
            pct = str(row["เปลี่ยนแปลง%"])
            up  = pct.startswith("+")
            for j, val in enumerate(hdrs, 1):
                c = ws8.cell(row=r_idx, column=j, value=row[val])
                c.border    = thin_border()
                c.alignment = Alignment(horizontal="center", vertical="center")
                if j == 1:
                    c.fill = PatternFill("solid", fgColor=h+"18"); c.font = Font(bold=True, color=h)
                elif j == 6:
                    c.fill = PatternFill("solid", fgColor="E8F5E9" if up else "FFEBEE")
                    c.font = Font(bold=True, color="2E7D32" if up else "C62828")
                else:
                    c.fill = PatternFill("solid", fgColor="F8F8F8" if r_idx%2==0 else "FFFFFF")
                    c.font = Font(color="333333", size=10)
            ws8.row_dimensions[r_idx].height = 20

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# ═══════════════════════════════════════════
# SESSION STATE
# ═══════════════════════════════════════════
for k, v in [("df",None),("cat_map",{}),("analyzed",False),("_file_id",""),
              ("df_prev",None),("cat_map_prev",{}),("analyzed_prev",False),("_file_id_prev",""),
              ("ai_status",""),("ai_remaining",0)]:
    if k not in st.session_state:
        st.session_state[k] = v

api_key = get_api_key()

# ═══════════════════════════════════════════
# HERO
# ═══════════════════════════════════════════
st.markdown("""
<div class="hero">
  <div class="hero-mascot">🛒</div>
  <div class="hero-title">CJ Smart Scan</div>
  <div class="hero-sub">📊 วิเคราะห์ประเภทสินค้าตาม PMA Groups + ยอดลูกค้าสะสมรายสาขา</div>
</div>
""", unsafe_allow_html=True)

if not api_key:
    st.warning("⚠️ ไม่พบ GEMINI_API_KEY — ใช้ rule-based เท่านั้น (ไม่มี AI)")

# ═══════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════
tab1, tab2 = st.tabs(["📦 เดือนปัจจุบัน", "📊 เปรียบเทียบกับเดือนก่อน"])

# ══════════════════════════════════
# TAB 1
# ══════════════════════════════════
with tab1:
    left, right = st.columns([1, 1.6], gap="large")

    with left:
        sec("อัปโหลดไฟล์ Excel", "📂")
        uploaded = st.file_uploader("เลือกไฟล์ Excel (sheet: ใบเสร็จ)", type=["xlsx"],
                                    key="up1", label_visibility="collapsed")
        if uploaded is not None:
            fid = f"{uploaded.name}_{uploaded.size}"
            if st.session_state._file_id != fid:
                try:
                    df_raw = load_excel(uploaded)
                    st.session_state.df       = df_raw
                    st.session_state.analyzed = False
                    st.session_state.cat_map  = {}
                    st.session_state._file_id = fid
                except Exception as e:
                    st.error(f"❌ {e}"); st.stop()

        if st.session_state.df is not None:
            n = int(st.session_state.df["ชื่อสินค้า"].notna().sum())
            b = int(st.session_state.df["รหัสสาขา"].dropna().nunique())
            st.success(f"✅ พบ **{n}** รายการ จาก **{b}** สาขา")

        if st.session_state.df is not None and not st.session_state.analyzed:
            if st.button("⚡ จำแนกสินค้าทันที", type="primary", use_container_width=True):
                prods = st.session_state.df["ชื่อสินค้า"].dropna().unique().tolist()
                st.session_state.cat_map  = classify_rule_only(prods)
                st.session_state.analyzed = True
                st.rerun()

        sec("หมวดหมู่ PMA", "🏷️")
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
            items      = df[df["ชื่อสินค้า"].notna()].copy()
            summary_df = make_summary(items)
            branch_df  = make_branch_summary(items)
            map_df     = pd.DataFrame(sorted(st.session_state.cat_map.items(), key=lambda x: x[1]),
                                      columns=["ชื่อสินค้า","ประเภทสินค้า"])
            cust_branch, cust_machine = get_customer_count(items)

            # Excel download
            has_prev = st.session_state.df_prev is not None and st.session_state.analyzed_prev
            sp = None
            if has_prev:
                df_p = st.session_state.df_prev.copy()
                df_p["ประเภทสินค้า"] = df_p["ชื่อสินค้า"].map(st.session_state.cat_map_prev)
                ip = df_p[df_p["ชื่อสินค้า"].notna()].copy()
                sp = make_summary(ip)

            excel_bytes = build_excel(df, summary_df, branch_df, map_df, items,
                                      cust_branch, cust_machine,
                                      df_prev=st.session_state.df_prev if has_prev else None,
                                      summary_prev=sp)
            lbl = "⬇️ ดาวน์โหลด Excel (พร้อมเปรียบเทียบ)" if has_prev else "⬇️ ดาวน์โหลด Excel"
            st.download_button(lbl, data=excel_bytes, file_name="CJ_PMA_classified.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)
            st.markdown("<br>", unsafe_allow_html=True)

            # Metrics
            m1, m2, m3, m4 = st.columns(4)
            total_cust = int(cust_branch["ยอดลูกค้าสะสม"].sum())
            for col, icon, num, lbl in [
                (m1, "🏪", b, "สาขา"),
                (m2, "📦", len(items), "รายการ"),
                (m3, "👥", f"{total_cust:,}", "ลูกค้าสะสม"),
                (m4, "🏷️", int(items["ประเภทสินค้า"].nunique()), "หมวด PMA"),
            ]:
                col.markdown(f'<div class="mcard"><div class="icon">{icon}</div>'
                             f'<div class="num">{num}</div><div class="lbl">{lbl}</div></div>',
                             unsafe_allow_html=True)

    # ── ยอดลูกค้าสะสม ──
    if st.session_state.analyzed and st.session_state.df is not None:
        df2 = st.session_state.df.copy()
        df2["ประเภทสินค้า"] = df2["ชื่อสินค้า"].map(st.session_state.cat_map)
        items2 = df2[df2["ชื่อสินค้า"].notna()].copy()
        cust_b2, cust_m2 = get_customer_count(items2)

        # AI button
        prods   = st.session_state.df["ชื่อสินค้า"].dropna().unique().tolist()
        need_ai = [p for p in prods if not rule_classify(p)]
        if st.session_state.get("ai_status") == "done":
            st.success("🎉 AI วิเคราะห์ครบทุกรายการแล้ว!")
            st.session_state.ai_status = ""
        if need_ai and api_key:
            c_info, c_btn = st.columns([2,1])
            c_info.info(f"มี **{len(need_ai)}** รายการยังไม่แน่ใจหมวดหมู่ — กด AI เพื่อเพิ่มความแม่นยำ")
            if c_btn.button("🤖 วิเคราะห์ด้วย Gemini AI", use_container_width=True):
                with st.spinner("🤖 กำลังวิเคราะห์..."):
                    new_map = classify_with_ai(prods, api_key, st.session_state.cat_map)
                st.session_state.cat_map = new_map
                still = [p for p in prods if not rule_classify(p) and new_map.get(p) not in CATEGORIES]
                st.session_state.ai_status = "done" if not still else "partial"
                st.rerun()
        elif not need_ai:
            st.success("✅ จำแนกครบทุกรายการแล้ว 🎉")

        st.markdown("<br>", unsafe_allow_html=True)
        sec("👥 ยอดลูกค้าสะสมต่อสาขา","")
        st.caption("ยอดลูกค้าสะสม = ผลรวมของ max(ลำดับใบเสร็จ) ต่อเครื่องคิดเงิน ของแต่ละสาขา")

        col_cust, col_mach = st.columns([1, 1], gap="large")
        with col_cust:
            top_cust = cust_b2.sort_values("ยอดลูกค้าสะสม", ascending=False).head(20)
            st.dataframe(top_cust.style
                         .format({"ยอดลูกค้าสะสม":"{:,.0f}"}),
                         use_container_width=True, hide_index=True, height=400)
        with col_mach:
            st.dataframe(cust_m2.sort_values(["รหัสสาขา","เครื่อง"])
                         .style.format({"ยอดลูกค้า":"{:,.0f}"}),
                         use_container_width=True, hide_index=True, height=400)

        # สัดส่วน PMA
        st.markdown("<br>", unsafe_allow_html=True)
        sec("📊 สัดส่วนตามหมวด PMA","")
        summary_df2 = make_summary(items2)
        col_bars, col_tbl = st.columns([1, 1], gap="large")
        with col_bars:
            max_cnt = int(summary_df2["จำนวนรายการ"].max()) or 1
            for _, row in summary_df2.iterrows():
                cat  = row["ประเภทสินค้า"]; cnt = int(row["จำนวนรายการ"])
                tot  = float(row["ยอดรวม"]) if pd.notna(row["ยอดรวม"]) else 0
                meta = CATEGORIES.get(cat, {"icon":"📦","color":"#9E9E9E"})
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
        with col_tbl:
            def hl(v):
                m = CATEGORIES.get(v, {"color":"#9E9E9E"})
                return f"background:{m['color']}15;color:{m['color']};font-weight:700"
            st.dataframe(
                summary_df2.style
                    .map(hl, subset=["ประเภทสินค้า"])
                    .format({"จำนวนรายการ":"{:,.0f}","ยอดรวม":"฿{:,.2f}"}),
                use_container_width=True, hide_index=True, height=380)

        # รายการสินค้า
        st.markdown("<br>", unsafe_allow_html=True)
        sec("📋 รายการสินค้าทั้งหมด","")
        fa, fb, fc = st.columns(3)
        bids = sorted(items2["รหัสสาขา"].dropna().unique().astype(int).tolist())
        with fa: sel_b = st.multiselect("🏪 สาขา", bids, default=bids, key="sb1", format_func=lambda x: f"สาขา {x}")
        with fb: sel_c = st.multiselect("🏷️ ประเภท PMA", list(CATEGORIES.keys()), default=list(CATEGORIES.keys()), key="sc1")
        with fc: srch  = st.text_input("🔍 ค้นหา", placeholder="ชื่อสินค้า...", key="sr1")
        filt = items2[items2["รหัสสาขา"].isin(sel_b) & items2["ประเภทสินค้า"].isin(sel_c)].copy()
        if srch: filt = filt[filt["ชื่อสินค้า"].str.contains(srch, na=False, case=False)]
        showcols = ["รหัสสาขา","วันที่","ชื่อสินค้า","ประเภทสินค้า","จำนวน","ราคาต่อหน่วย","ยอดรวมสินค้า"]
        st.dataframe(filt[[c for c in showcols if c in filt.columns]]
                     .style.map(hl, subset=["ประเภทสินค้า"])
                     .format({"จำนวน":"{:.0f}","ราคาต่อหน่วย":"฿{:,.2f}","ยอดรวมสินค้า":"฿{:,.2f}"}),
                     use_container_width=True, hide_index=True, height=480)
        st.caption(f"แสดง **{len(filt):,}** จาก **{len(items2):,}** รายการ")

        # สรุปรายสาขา
        st.markdown("<br>", unsafe_allow_html=True)
        sec("🏪 สรุปรายสาขา","")
        bdf3  = make_branch_summary(items2)
        bids3 = sorted(items2["รหัสสาขา"].dropna().unique().astype(int).tolist())
        tabs_b = st.tabs([f"🏪 {b}" for b in bids3])
        for tab, b in zip(tabs_b, bids3):
            with tab:
                bd = bdf3[bdf3["รหัสสาขา"] == b]
                c1, c2 = st.columns([1.4, 1])
                with c1:
                    st.bar_chart(bd.set_index("ประเภทสินค้า")["จำนวนรายการ"], color="#FF4D6D", height=300)
                with c2:
                    # ยอดลูกค้าของสาขานี้
                    cust_this = cust_b2[cust_b2["รหัสสาขา"] == b]
                    if len(cust_this):
                        st.metric("👥 ยอดลูกค้าสะสม", f'{int(cust_this["ยอดลูกค้าสะสม"].iloc[0]):,} transaction')
                    st.dataframe(bd[["ประเภทสินค้า","จำนวนรายการ","ยอดรวม"]]
                                 .style.format({"ยอดรวม":"฿{:,.2f}"}),
                                 use_container_width=True, hide_index=True, height=250)

# ══════════════════════════════════
# TAB 2 — เปรียบเทียบ
# ══════════════════════════════════
with tab2:
    st.markdown("### 📊 เปรียบเทียบกับเดือนก่อน")
    col_up, col_info = st.columns([1, 1.6], gap="large")

    with col_up:
        sec("อัปโหลดไฟล์เดือนก่อน","📂")
        uploaded_prev = st.file_uploader("เลือกไฟล์ Excel เดือนก่อน", type=["xlsx"],
                                         key="up2", label_visibility="collapsed")
        if uploaded_prev is not None:
            fid2 = f"{uploaded_prev.name}_{uploaded_prev.size}"
            if st.session_state._file_id_prev != fid2:
                try:
                    df_p = load_excel(uploaded_prev)
                    st.session_state.df_prev       = df_p
                    st.session_state.analyzed_prev = False
                    st.session_state.cat_map_prev  = {}
                    st.session_state._file_id_prev = fid2
                except Exception as e:
                    st.error(f"❌ {e}"); st.stop()

        if st.session_state.df_prev is not None:
            n2 = int(st.session_state.df_prev["ชื่อสินค้า"].notna().sum())
            b2 = int(st.session_state.df_prev["รหัสสาขา"].dropna().nunique())
            st.success(f"✅ พบ **{n2}** รายการ จาก **{b2}** สาขา")

        if st.session_state.df_prev is not None and not st.session_state.analyzed_prev:
            if st.button("⚡ จำแนกไฟล์เดือนก่อน", type="primary", use_container_width=True, key="btn_prev"):
                prods2 = st.session_state.df_prev["ชื่อสินค้า"].dropna().unique().tolist()
                st.session_state.cat_map_prev  = classify_rule_only(prods2)
                st.session_state.analyzed_prev = True
                st.rerun()

    with col_info:
        if st.session_state.analyzed and st.session_state.analyzed_prev:
            df_cur = st.session_state.df.copy()
            df_cur["ประเภทสินค้า"] = df_cur["ชื่อสินค้า"].map(st.session_state.cat_map)
            items_cur = df_cur[df_cur["ชื่อสินค้า"].notna()].copy()
            sum_cur   = make_summary(items_cur)
            cust_cur, _ = get_customer_count(items_cur)

            df_prv = st.session_state.df_prev.copy()
            df_prv["ประเภทสินค้า"] = df_prv["ชื่อสินค้า"].map(st.session_state.cat_map_prev)
            items_prv = df_prv[df_prv["ชื่อสินค้า"].notna()].copy()
            sum_prv   = make_summary(items_prv)
            cust_prv, _ = get_customer_count(items_prv)

            lbl_cur = items_cur["เดือน"].iloc[0] if len(items_cur) else "ปัจจุบัน"
            lbl_prv = items_prv["เดือน"].iloc[0] if len(items_prv) else "ก่อนหน้า"

            tot_cust_cur = int(cust_cur["ยอดลูกค้าสะสม"].sum())
            tot_cust_prv = int(cust_prv["ยอดลูกค้าสะสม"].sum())
            chg_pct = ((tot_cust_cur - tot_cust_prv) / tot_cust_prv * 100) if tot_cust_prv > 0 else 0
            arrow   = "📈" if chg_pct >= 0 else "📉"
            color   = "#2E7D32" if chg_pct >= 0 else "#C62828"

            m1, m2, m3 = st.columns(3)
            m1.markdown(f'<div class="mcard"><div class="icon">👥</div>'
                        f'<div class="num">{tot_cust_prv:,}</div>'
                        f'<div class="lbl">ลูกค้า {lbl_prv}</div></div>', unsafe_allow_html=True)
            m2.markdown(f'<div class="mcard"><div class="icon">👥</div>'
                        f'<div class="num">{tot_cust_cur:,}</div>'
                        f'<div class="lbl">ลูกค้า {lbl_cur}</div></div>', unsafe_allow_html=True)
            m3.markdown(f'<div class="mcard"><div class="icon">{arrow}</div>'
                        f'<div class="num" style="color:{color}">{chg_pct:+.1f}%</div>'
                        f'<div class="lbl">เปลี่ยนแปลงลูกค้า</div></div>', unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            sec("เปรียบเทียบลูกค้าสะสมต่อสาขา","👥")
            br_mrg = cust_cur.merge(cust_prv, on="รหัสสาขา", how="outer", suffixes=("_cur","_prv")).fillna(0)
            br_mrg["เปลี่ยนแปลง"] = br_mrg["ยอดลูกค้าสะสม_cur"] - br_mrg["ยอดลูกค้าสะสม_prv"]
            br_mrg["เปลี่ยนแปลง%"] = br_mrg.apply(
                lambda r: f'{(r["เปลี่ยนแปลง"]/r["ยอดลูกค้าสะสม_prv"]*100):+.1f}%'
                if r["ยอดลูกค้าสะสม_prv"] > 0 else "ใหม่", axis=1)
            br_mrg.columns = ["รหัสสาขา", f"ลูกค้า_{lbl_prv}", f"ลูกค้า_{lbl_cur}", "เปลี่ยนแปลง","เปลี่ยนแปลง%"]
            st.dataframe(br_mrg.sort_values(f"ลูกค้า_{lbl_cur}", ascending=False)
                         .style.format({f"ลูกค้า_{lbl_prv}":"{:,.0f}", f"ลูกค้า_{lbl_cur}":"{:,.0f}", "เปลี่ยนแปลง":"{:+,.0f}"}),
                         use_container_width=True, hide_index=True, height=400)

            st.markdown("<br>", unsafe_allow_html=True)
            sec("เปรียบเทียบตามประเภท PMA","📊")
            merged = sum_cur.rename(columns={"จำนวนรายการ":"จำนวน_ปัจจุบัน","ยอดรวม":"ยอด_ปัจจุบัน"}).merge(
                sum_prv.rename(columns={"จำนวนรายการ":"จำนวน_เดือนก่อน","ยอดรวม":"ยอด_เดือนก่อน"}),
                on="ประเภทสินค้า", how="outer").fillna(0)
            merged["เปลี่ยนแปลง%"] = merged.apply(
                lambda r: f'{((r["ยอด_ปัจจุบัน"]-r["ยอด_เดือนก่อน"])/r["ยอด_เดือนก่อน"]*100):+.1f}%'
                if r["ยอด_เดือนก่อน"] > 0 else ("ใหม่" if r["ยอด_ปัจจุบัน"] > 0 else "-"), axis=1)
            for _, row in merged.iterrows():
                cat  = row["ประเภทสินค้า"]
                meta = CATEGORIES.get(cat, {"icon":"📦","color":"#9E9E9E"})
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
        else:
            st.info("📌 อัปโหลดและจำแนกไฟล์ทั้งสองเดือนก่อนเพื่อดูผลเปรียบเทียบ")

st.markdown('<div class="footer">🛒 CJ Smart Scan · PMA Groups · Powered by Gemini AI</div>',
            unsafe_allow_html=True)
