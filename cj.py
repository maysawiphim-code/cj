import streamlit as st
import pandas as pd
import json, re, io, urllib.request
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="CJ Smart Scan", page_icon="🛒", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;500;600;700;800&display=swap');
@import url('https://fonts.googleapis.com/css2?family=Fredoka+One&display=swap');
*,[class*="css"]{font-family:'Sarabun',sans-serif!important}
#MainMenu,header,footer{visibility:hidden}
.stApp{background:linear-gradient(160deg,#FFF8F0 0%,#FFF0F5 50%,#F0F4FF 100%)}
.hero{background:linear-gradient(135deg,#FF6B6B 0%,#FF4D6D 40%,#C9184A 100%);
  border-radius:24px;padding:32px 40px 28px;margin-bottom:24px;position:relative;
  overflow:hidden;box-shadow:0 8px 32px rgba(201,24,74,.25)}
.hero-title{font-family:'Fredoka One','Sarabun',sans-serif!important;
  font-size:2.4rem;font-weight:900;color:white;margin:0;text-shadow:0 2px 8px rgba(0,0,0,.2)}
.hero-sub{color:rgba(255,255,255,.88);font-size:1rem;margin:8px 0 0}
.stDownloadButton>button,.stButton>button[kind="primary"]{
  background:linear-gradient(135deg,#FF6B6B,#C9184A)!important;color:white!important;
  border:none!important;border-radius:12px!important;padding:12px 24px!important;font-weight:700!important}
</style>
""", unsafe_allow_html=True)

# ========================= CATEGORIES =========================
CATEGORIES = {
    "Fresh Food": {"icon":"🥩","color":"#E53935","hex":"E53935"},
    "Non Food": {"icon":"🏠","color":"#1565C0","hex":"1565C0"},
    "Packaged Beverage": {"icon":"🥤","color":"#00897B","hex":"00897B"},
    "Processed Food": {"icon":"🍜","color":"#F57C00","hex":"F57C00"},
    "Special Business": {"icon":"⭐","color":"#7B1FA2","hex":"7B1FA2"},
}

# ========================= RULE_KW =========================
RULE_KW = {
    "Fresh Food": {
        "keywords": ["bao","bac","แซนด์วิช","ไส้กรอก","ขนมปัง","เบเกอรี","sandwich","sausage","ข้าวกล่อง","สลัด","salad","นม pasteur","โยเกิร์ต","ไก่","หมู","เนื้อ","ทูน่า","ปลา","กริล","อบ","ปิ้ง","โรตี","ครัวซอง","โดนัท","เค้ก","พาย","fresh","bakery","counter","warm","meal"],
        "confidence": 0.92
    },
    "Non Food": {
        "keywords": ["ยาสีฟัน","แชมพู","สบู่","ครีม","ผ้าอนามัย","ทิชชู","ถุงขยะ","ผงซักฟอก","บุหรี่","cigarette","ไฟแช็ก","ถ่าน","แปรง","โลชั่น","ปากกา","สมุด","tissue","soap","shampoo","toothpaste","houseware","stationery"],
        "confidence": 0.90
    },
    "Packaged Beverage": {
        "keywords": ["น้ำดื่ม","น้ำแร่","โค้ก","เป๊ปซี่","สไปรท์","คาราบาว","เรดบูล","ไมโล","นมuht","uht","เบียร์","beer","วิสกี้","สุรา","ไอศกรีม","ice cream","ชาเขียว","โอเลี้ยง","energy","soft drink","โซดา","csd","น้ำอัดลม"],
        "confidence": 0.88
    },
    "Processed Food": {
        "keywords": ["มาม่า","บะหมี่","instant","ขนมกรุบ","มันฝรั่ง","ลูกอม","ช็อกโกแลต","ถั่ว","snack","canned","กระป๋อง","ขนม","biscuit","cookie","pasta"],
        "confidence": 0.85
    },
    "Special Business": {
        "keywords": ["เติมเงิน","จ่ายบิล","ส่วนลด","ยา","vitamin","วิตามิน","bellinee","kudsan","7-service","คูปอง","บริการ","discount","social welfare"],
        "confidence": 0.95
    },
}

SYSTEM_PROMPT = """คุณคือผู้เชี่ยวชาญการจำแนกสินค้า CJ Express ตาม PMA 5 กลุ่มอย่างเคร่งครัด

1. Fresh Food → อาหารสด เบเกอรี่สด แซนด์วิช ไส้กรอก ข้าวกล่อง สลัด นมพาสเจอร์ไรส์
2. Non Food → ของใช้ส่วนตัว ของใช้ในบ้าน บุหรี่ เครื่องเขียน ทิชชู ผ้าอนามัย
3. Packaged Beverage → น้ำดื่ม น้ำอัดลม เบียร์ สุรา ไอศกรีม นม UHT
4. Processed Food → อาหารแปรรูป บะหมี่ ขนมขบเคี้ยว ช็อกโกแลต อาหารกระป๋อง
5. Special Business → ยา วิตามิน บริการเติมเงิน จ่ายบิล Bellinee Kudsan

ตอบเป็น JSON เท่านั้น: {"ชื่อสินค้า": "ชื่อกลุ่ม", ...}"""

# ========================= FUNCTIONS =========================
def get_api_key():
    import os
    try:
        k = st.secrets["GEMINI_API_KEY"]
        if k and str(k).strip(): return str(k).strip()
    except: pass
    return os.environ.get("GEMINI_API_KEY", "").strip()

def rule_classify_with_confidence(name: str):
    n = str(name).lower().strip()
    best_cat = None
    best_score = 0
    best_conf = 0.0
    for cat, data in RULE_KW.items():
        score = sum(1 for kw in data["keywords"] if kw.lower() in n)
        if score > best_score:
            best_score = score
            best_cat = cat
            best_conf = data["confidence"]
    return best_cat, best_conf if best_score > 0 else 0.0

def gemini_classify(products, api_key):
    if not products:
        return {}
    product_list = "\n".join(f"- {p}" for p in products)
    prompt = SYSTEM_PROMPT + f"\n\nสินค้าที่ต้องจำแนก:\n{product_list}"
    
    payload = json.dumps({
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.05, "maxOutputTokens": 4000, "topP": 0.9}
    }, ensure_ascii=False).encode("utf-8")

    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={api_key}"
    try:
        req = urllib.request.Request(url, data=payload, headers={"Content-Type": "application/json; charset=utf-8"})
        with urllib.request.urlopen(req, timeout=90) as r:
            data = json.loads(r.read().decode("utf-8"))
        raw_text = data["candidates"][0]["content"]["parts"][0]["text"]
        # แก้ไขตรงนี้ให้ถูกต้อง
        raw = re.sub(r"```json|```", "", raw_text).strip()
        result = json.loads(raw)
        valid = set(CATEGORIES.keys())
        return {k: (v if v in valid else "Processed Food") for k, v in result.items()}
    except Exception as e:
        st.warning(f"Gemini Error: {str(e)[:150]}")
        return {}

def post_validation(cat_map):
    validated = {}
    for product, category in cat_map.items():
        p = str(product).lower()
        if any(x in p for x in ["ไอศกรีม", "ice cream"]):
            validated[product] = "Packaged Beverage"
        elif any(x in p for x in ["ขนมปัง", "เบเกอรี", "bakery", "โดนัท", "เค้ก", "โรตี", "ครัวซอง"]):
            validated[product] = "Fresh Food"
        elif any(x in p for x in ["ยา", "vitamin", "วิตามิน", "bellinee", "kudsan"]):
            validated[product] = "Special Business"
        else:
            validated[product] = category
    return validated

def smart_hybrid_classify(products, api_key, existing_map=None):
    if existing_map is None:
        existing_map = {}
    result = dict(existing_map)
    to_ai = []
    
    for p in products:
        if p in result and result[p]:
            continue
        cat, conf = rule_classify_with_confidence(p)
        if conf >= 0.75:
            result[p] = cat
        else:
            to_ai.append(p)
    
    if to_ai and api_key:
        st.info(f"Rule-based จัดได้ {len(products)-len(to_ai)} รายการ | ใช้ AI กับ **{len(to_ai)}** รายการ")
        batches = [to_ai[i:i+25] for i in range(0, len(to_ai), 25)]
        bar = st.progress(0)
        status = st.empty()
        
        for i, batch in enumerate(batches):
            status.markdown(f"🤖 Gemini AI กำลังวิเคราะห์... ({i+1}/{len(batches)})")
            ai_result = gemini_classify(batch, api_key)
            result.update(ai_result)
            bar.progress((i + 1) / len(batches))
        
        bar.empty()
        status.empty()
    
    result = post_validation(result)
    return result

# ========================= CORE =========================
def load_excel(file):
    df = pd.read_excel(file, sheet_name="ใบเสร็จ")
    drop = [c for c in df.columns if "หมวด" in str(c) or "ประเภท" in str(c) or "category" in str(c).lower()]
    if drop:
        df = df.drop(columns=drop)
    df["วันที่_dt"] = pd.to_datetime(df["วันที่"], format="%d/%m/%Y", errors="coerce")
    df["เดือน"] = df["วันที่_dt"].dt.strftime("%b %Y")
    return df

def parse_receipt_no(df):
    def extract(r):
        m = re.search(r'(N\d+)-(\d+)', str(r))
        if m:
            return m.group(1), int(m.group(2))
        return None, None
    df = df.copy()
    parsed = df["เลขที่ใบเสร็จ"].apply(extract)
    df["เครื่อง"] = parsed.apply(lambda x: x[0])
    df["ยอดลูกค้า"] = parsed.apply(lambda x: x[1])
    return df

def get_customer_accumulate(df):
    df = parse_receipt_no(df)
    df_valid = df.dropna(subset=["เครื่อง", "ยอดลูกค้า"])
    if df_valid.empty:
        return pd.DataFrame(columns=["รหัสสาขา", "เครื่อง", "ยอดลูกค้าสะสม"])
    machine_max = df_valid.groupby(["รหัสสาขา", "เครื่อง"])["ยอดลูกค้า"].max().reset_index()
    machine_max.columns = ["รหัสสาขา", "เครื่อง", "ยอดลูกค้าสะสม"]
    return machine_max

def make_branch_customer_summary(df):
    mc = get_customer_accumulate(df)
    branch_total = mc.groupby("รหัสสาขา")["ยอดลูกค้าสะสม"].sum().reset_index()
    branch_total.columns = ["รหัสสาขา", "ยอดลูกค้าสะสมรวม"]
    branch_total = branch_total.sort_values("ยอดลูกค้าสะสมรวม", ascending=False)
    return mc, branch_total

def make_summary(items):
    grp = items.groupby("ประเภทสินค้า")
    df = pd.concat([grp["ชื่อสินค้า"].count(), grp["ยอดรวมสินค้า"].sum()], axis=1)
    df.columns = ["จำนวนรายการ", "ยอดรวม"]
    return df.reset_index().sort_values("จำนวนรายการ", ascending=False)

def build_excel(df, summary, items, mc_df, branch_tot_df, map_df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="ข้อมูลทั้งหมด", index=False, startrow=2)
        summary.to_excel(writer, sheet_name="สรุปตามประเภท", index=False, startrow=2)
        branch_tot_df.to_excel(writer, sheet_name="ยอดลูกค้าสะสมรายสาขา", index=False, startrow=2)
        mc_df.to_excel(writer, sheet_name="ยอดลูกค้าแยกเครื่อง", index=False, startrow=2)
        map_df.to_excel(writer, sheet_name="mapping สินค้า", index=False, startrow=2)
    buf.seek(0)
    out = io.BytesIO()
    wb = load_workbook(buf)
    wb.save(out)
    return out.getvalue()

# ========================= SESSION STATE =========================
for k, v in [("df", None), ("cat_map", {}), ("analyzed", False), ("_fid", ""),
             ("df_prev", None), ("cat_map_prev", {}), ("analyzed_prev", False), ("_fid_prev", "")]:
    if k not in st.session_state:
        st.session_state[k] = v

api_key = get_api_key()

# ========================= UI =========================
st.markdown("""
<div class="hero">
  <div class="hero-title">CJ Smart Scan</div>
  <div class="hero-sub">📊 วิเคราะห์ยอดลูกค้าสะสม & จำแนกสินค้า PMA 5 กลุ่ม (Smart Hybrid)</div>
</div>
""", unsafe_allow_html=True)

if not api_key:
    st.error("⚠️ ไม่พบ GEMINI_API_KEY — กรุณาตั้งค่าใน Streamlit Secrets")
    st.stop()

tab1, tab2 = st.tabs(["📦 วิเคราะห์ไฟล์", "📊 เปรียบเทียบเดือน"])

with tab1:
    left, right = st.columns([1, 1.6], gap="large")
    with left:
        uploaded = st.file_uploader("เลือกไฟล์ Excel", type=["xlsx"], key="up1")
        if uploaded is not None:
            fid = f"{uploaded.name}_{uploaded.size}"
            if st.session_state._fid != fid:
                try:
                    df_raw = load_excel(uploaded)
                    st.session_state.df = df_raw
                    st.session_state.analyzed = False
                    st.session_state.cat_map = {}
                    st.session_state._fid = fid
                    st.success("✅ โหลดไฟล์สำเร็จ")
                except Exception as e:
                    st.error(f"❌ {e}")

        if st.session_state.df is not None and not st.session_state.analyzed:
            if st.button("⚡ จำแนกสินค้า (Smart Hybrid)", type="primary", use_container_width=True):
                prods = st.session_state.df["ชื่อสินค้า"].dropna().unique().tolist()
                with st.spinner("กำลังจำแนกด้วย Smart Hybrid Logic..."):
                    st.session_state.cat_map = smart_hybrid_classify(prods, api_key, st.session_state.cat_map)
                    st.session_state.analyzed = True
                st.success("✅ จำแนกสินค้าเสร็จสิ้น!")
                st.rerun()

    with right:
        if st.session_state.analyzed and st.session_state.df is not None:
            df = st.session_state.df.copy()
            df["ประเภทสินค้า"] = df["ชื่อสินค้า"].map(st.session_state.cat_map)
            items = df[df["ชื่อสินค้า"].notna()].copy()
            mc_df, branch_tot_df = make_branch_customer_summary(df)
            summary_df = make_summary(items)
            map_df = pd.DataFrame(sorted(st.session_state.cat_map.items()), columns=["ชื่อสินค้า", "ประเภทสินค้า"])

            excel_bytes = build_excel(df, summary_df, items, mc_df, branch_tot_df, map_df)
            st.download_button("⬇️ ดาวน์โหลด Excel", 
                               data=excel_bytes, 
                               file_name="CJ_SmartScan.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)

st.caption("💡 Smart Hybrid Logic (Rule-based + Gemini AI + Post Validation)")
