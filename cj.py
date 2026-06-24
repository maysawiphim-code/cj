import streamlit as st
import pandas as pd
import json, re, io, urllib.request

st.set_page_config(page_title="CJ Smart Scan 🏪", page_icon="🛒", layout="wide")

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
    "Bao Cafe":         {"icon":"☕","color":"#C2185B"},
    "อาหารพร้อมทาน":   {"icon":"🍱","color":"#F57C00"},
    "เครื่องดื่ม":      {"icon":"🥤","color":"#1976D2"},
    "ขนมขบเคี้ยว":     {"icon":"🍿","color":"#7B1FA2"},
    "ของใช้ส่วนตัว":   {"icon":"🧴","color":"#00897B"},
    "ของใช้ในบ้าน":    {"icon":"🏠","color":"#558B2F"},
    "สินค้าเบ็ดเตล็ด": {"icon":"🔋","color":"#6D4C41"},
    "บริการและอื่นๆ":  {"icon":"📱","color":"#455A64"},
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
    key = os.environ.get("ANTHROPIC_API_KEY","")
    return key.strip() if key else ""

def claude_classify(products, api_key):
    product_list = "\n".join(f"- {p}" for p in products)
    payload = json.dumps({
        "model": "claude-sonnet-4-6",
        "max_tokens": 2000,
        "system": SYSTEM_PROMPT,
        "messages": [{"role":"user","content":f"วิเคราะห์ประเภทสินค้าจากชื่อต่อไปนี้:\n{product_list}"}],
    }).encode()
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages", data=payload,
        headers={"Content-Type":"application/json","x-api-key":api_key,"anthropic-version":"2023-06-01"},
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
        status.markdown(f"🤖 กำลังวิเคราะห์ชื่อสินค้า **{min((i+1)*batch_size,len(products))}/{len(products)}** รายการ...")
        try:
            result.update(claude_classify(batch, api_key))
        except Exception as e:
            st.warning(f"Batch {i+1} error: {e}")
            for p in batch: result[p] = "สินค้าเบ็ดเตล็ด"
        bar.progress((i+1)/len(batches))
    bar.empty(); status.empty()
    return result

def build_excel(df, summary, branch_df, map_df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="ข้อมูลทั้งหมด", index=False)
        summary.to_excel(w, sheet_name="สรุปตามประเภท", index=False)
        branch_df.to_excel(w, sheet_name="สรุปตามสาขา", index=False)
        map_df.to_excel(w, sheet_name="mapping สินค้า", index=False)
    return buf.getvalue()

def make_summary(items):
    grp = items.groupby("ประเภทสินค้า")
    cnt = grp["ชื่อสินค้า"].count()
    tot = grp["ยอดรวมสินค้า"].sum()
    df = pd.concat([cnt, tot], axis=1)
    df.columns = ["จำนวนรายการ","ยอดรวม"]
    return df.reset_index().sort_values("จำนวนรายการ", ascending=False)

def make_branch_summary(items):
    grp = items.groupby(["รหัสสาขา","ประเภทสินค้า"])
    cnt = grp["ชื่อสินค้า"].count()
    tot = grp["ยอดรวมสินค้า"].sum()
    df = pd.concat([cnt, tot], axis=1)
    df.columns = ["จำนวนรายการ","ยอดรวม"]
    return df.reset_index().sort_values(["รหัสสาขา","จำนวนรายการ"], ascending=[True,False])

def sec(title, icon=""):
    st.markdown(f'<div class="sec-title"><div class="dot"></div>{icon} {title}</div>', unsafe_allow_html=True)

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
        try:
            st.write("Keys:", list(st.secrets.keys()))
        except Exception as e:
            st.write("อ่าน secrets ไม่ได้:", e)
    st.info("ไปที่ App Settings → Secrets แล้วใส่:\n```\nANTHROPIC_API_KEY = \"sk-ant-...\"\n```")
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
            if st.session_state.df is None or len(st.session_state.df) != len(df_raw):
                st.session_state.df = df_raw
                st.session_state.analyzed = False
                st.session_state.cat_map = {}
            n_items = int(df_raw["ชื่อสินค้า"].notna().sum())
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
                f'<div style="display:flex;align-items:center;gap:10px;padding:6px 0">'
                f'<span style="font-size:1.4rem">{meta["icon"]}</span>'
                f'<span style="background:{meta["color"]};color:white;padding:3px 12px;'
                f'border-radius:99px;font-size:.8rem;font-weight:600">{name}</span>'
                f'</div>', unsafe_allow_html=True)

with right:
    if st.session_state.analyzed and st.session_state.df is not None:
        df = st.session_state.df.copy()
        df["ประเภทสินค้า"] = df["ชื่อสินค้า"].map(st.session_state.cat_map)
        items = df[df["ชื่อสินค้า"].notna()].copy()

        summary_df  = make_summary(items)
        branch_df   = make_branch_summary(items)
        map_df      = pd.DataFrame(
            sorted(st.session_state.cat_map.items(), key=lambda x: x[1]),
            columns=["ชื่อสินค้า","ประเภทสินค้า"]
        )
        excel_bytes = build_excel(df, summary_df, branch_df, map_df)

        st.download_button(
            "⬇️  ดาวน์โหลด Excel สรุปผลทั้งหมด",
            data=excel_bytes,
            file_name="CJ_product_classified.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.markdown("<br>", unsafe_allow_html=True)

        # Metrics
        total_sales = float(items["ยอดรวมสินค้า"].sum())
        m1,m2,m3,m4 = st.columns(4)
        for col,icon,num,lbl in [
            (m1,"🏪", int(items["รหัสสาขา"].nunique()), "สาขา"),
            (m2,"📦", len(items), "รายการ"),
            (m3,"🏷️", int(items["ประเภทสินค้า"].nunique()), "หมวดหมู่"),
            (m4,"💰", f"฿{total_sales:,.0f}", "ยอดรวม"),
        ]:
            col.markdown(f'<div class="mcard"><div class="icon">{icon}</div>'
                         f'<div class="num">{num}</div><div class="lbl">{lbl}</div></div>',
                         unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        sec("สัดส่วนตามหมวดหมู่ (วิเคราะห์จากชื่อสินค้า)", "📊")
        max_cnt = int(summary_df["จำนวนรายการ"].max()) or 1
        for _, row in summary_df.iterrows():
            cat  = row["ประเภทสินค้า"]
            cnt  = int(row["จำนวนรายการ"])
            tot  = float(row["ยอดรวม"]) if pd.notna(row["ยอดรวม"]) else 0
            meta = CATEGORIES.get(cat, {"icon":"📦","color":"#9E9E9E"})
            pct  = cnt / max_cnt * 100
            st.markdown(f"""
            <div class="cat-card">
              <div class="cat-icon">{meta["icon"]}</div>
              <div style="flex:1">
                <div class="cat-name">{cat}</div>
                <div class="cat-count">฿{tot:,.0f}</div>
                <div class="bar-wrap" style="margin-top:6px">
                  <div class="bar-fill" style="width:{pct}%;background:{meta['color']}"></div>
                </div>
              </div>
              <div class="cat-num">{cnt}</div>
            </div>""", unsafe_allow_html=True)

# ─── DETAIL TABLE ───
if st.session_state.analyzed and st.session_state.df is not None:
    df = st.session_state.df.copy()
    df["ประเภทสินค้า"] = df["ชื่อสินค้า"].map(st.session_state.cat_map)
    items = df[df["ชื่อสินค้า"].notna()].copy()

    st.markdown("<br>", unsafe_allow_html=True)
    sec("รายการสินค้าทั้งหมด (จำแนกจากชื่อสินค้า)", "📋")

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

    cols = ["รหัสสาขา","วันที่","ชื่อสินค้า","ประเภทสินค้า","จำนวน","ราคาต่อหน่วย","ยอดรวมสินค้า"]
    def highlight_cat(val):
        meta = CATEGORIES.get(val, {"color":"#9E9E9E"})
        return f"background-color:{meta['color']}22;color:{meta['color']};font-weight:600"

    styled = filtered[cols].style\
        .map(highlight_cat, subset=["ประเภทสินค้า"])\
        .format({"จำนวน":"{:.0f}","ราคาต่อหน่วย":"฿{:,.2f}","ยอดรวมสินค้า":"฿{:,.2f}"})

    st.dataframe(styled, use_container_width=True, hide_index=True, height=420)
    st.caption(f"แสดง **{len(filtered):,}** จาก **{len(items):,}** รายการ")

    # Branch tabs
    st.markdown("<br>", unsafe_allow_html=True)
    sec("สรุปรายสาขา", "🏪")
    branch_df2 = make_branch_summary(items)
    tabs = st.tabs([f"🏪 สาขา {b}" for b in branch_ids])
    for tab, b in zip(tabs, branch_ids):
        with tab:
            bd = branch_df2[branch_df2["รหัสสาขา"] == b].copy()
            c1, c2 = st.columns([1.2, 1])
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
