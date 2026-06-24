import streamlit as st
import pandas as pd
import json, re, io, urllib.request

# ─────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────
st.set_page_config(
    page_title="CJ Smart Scan 🏪",
    page_icon="🛒",
    layout="wide",
)

# ─────────────────────────────────────────
# GLOBAL CSS
# ─────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;500;600;700;800&display=swap');
@import url('https://fonts.googleapis.com/css2?family=Fredoka+One&display=swap');

*, html, body, [class*="css"] {
    font-family: 'Sarabun', sans-serif !important;
}

/* ── hide default header ── */
#MainMenu, header, footer { visibility: hidden; }

/* ── page background ── */
.stApp {
    background: linear-gradient(160deg, #FFF8F0 0%, #FFF0F5 50%, #F0F4FF 100%);
    min-height: 100vh;
}

/* ── hero banner ── */
.hero {
    background: linear-gradient(135deg, #FF6B6B 0%, #FF4D6D 40%, #C9184A 100%);
    border-radius: 24px;
    padding: 36px 40px 32px;
    margin-bottom: 32px;
    position: relative;
    overflow: hidden;
    box-shadow: 0 8px 32px rgba(201,24,74,.25);
}
.hero::before {
    content: '';
    position: absolute; inset: 0;
    background: url("data:image/svg+xml,%3Csvg width='60' height='60' viewBox='0 0 60 60' xmlns='http://www.w3.org/2000/svg'%3E%3Cg fill='none' fill-rule='evenodd'%3E%3Cg fill='%23ffffff' fill-opacity='0.06'%3E%3Ccircle cx='30' cy='30' r='20'/%3E%3C/g%3E%3C/g%3E%3C/svg%3E");
}
.hero-title {
    font-family: 'Fredoka One', 'Sarabun', sans-serif !important;
    font-size: 2.6rem; font-weight: 900;
    color: white; margin: 0; line-height: 1.2;
    text-shadow: 0 2px 8px rgba(0,0,0,.2);
}
.hero-sub { color: rgba(255,255,255,.88); font-size: 1.05rem; margin: 10px 0 0; }
.hero-mascot {
    position: absolute; right: 40px; top: 50%; transform: translateY(-50%);
    font-size: 7rem; line-height: 1; filter: drop-shadow(0 4px 12px rgba(0,0,0,.15));
}

/* ── section title ── */
.sec-title {
    display: flex; align-items: center; gap: 10px;
    font-size: 1.15rem; font-weight: 700; color: #C9184A;
    margin: 28px 0 14px;
}
.sec-title .dot {
    width: 6px; height: 24px; background: #FF4D6D;
    border-radius: 3px; flex-shrink: 0;
}

/* ── upload card ── */
.upload-card {
    background: white;
    border: 2.5px dashed #FFB3C1;
    border-radius: 20px;
    padding: 32px;
    text-align: center;
    transition: border-color .2s;
}
.upload-card:hover { border-color: #FF4D6D; }

/* ── metric card ── */
.mcard {
    background: white;
    border-radius: 18px;
    padding: 22px 18px;
    text-align: center;
    box-shadow: 0 4px 16px rgba(0,0,0,.07);
    transition: transform .15s;
}
.mcard:hover { transform: translateY(-3px); }
.mcard .icon { font-size: 2.4rem; line-height: 1; margin-bottom: 6px; }
.mcard .num  { font-size: 2rem; font-weight: 800; color: #C9184A; line-height: 1.1; }
.mcard .lbl  { font-size: .82rem; color: #888; margin-top: 4px; font-weight: 500; }

/* ── category pill ── */
.cat-pill {
    display: inline-flex; align-items: center; gap: 5px;
    padding: 4px 12px; border-radius: 99px;
    font-size: .8rem; font-weight: 600; color: white;
    white-space: nowrap;
}

/* ── cat summary card ── */
.cat-card {
    background: white;
    border-radius: 16px;
    padding: 16px 20px;
    margin-bottom: 10px;
    display: flex; align-items: center; gap: 14px;
    box-shadow: 0 2px 10px rgba(0,0,0,.06);
}
.cat-card .cat-icon { font-size: 2rem; line-height: 1; }
.cat-card .cat-name { font-weight: 700; font-size: .95rem; color: #333; }
.cat-card .cat-count { font-size: .82rem; color: #888; }
.cat-card .bar-wrap  { flex: 1; background: #f5f5f5; border-radius: 99px; height: 8px; overflow: hidden; }
.cat-card .bar-fill  { height: 100%; border-radius: 99px; }
.cat-card .cat-num   { font-weight: 800; font-size: 1.1rem; color: #C9184A; min-width: 36px; text-align: right; }

/* ── branch tab style ── */
div[data-baseweb="tab-list"] {
    background: white; border-radius: 12px; padding: 4px;
    box-shadow: 0 2px 8px rgba(0,0,0,.06);
    gap: 4px;
}
button[data-baseweb="tab"] {
    border-radius: 10px !important;
    font-weight: 600 !important;
}

/* ── download button ── */
.stDownloadButton > button {
    background: linear-gradient(135deg, #FF6B6B, #C9184A) !important;
    color: white !important; border: none !important;
    border-radius: 14px !important; padding: 14px 28px !important;
    font-size: 1rem !important; font-weight: 700 !important;
    box-shadow: 0 4px 16px rgba(201,24,74,.3) !important;
    transition: all .2s !important;
}
.stDownloadButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 6px 20px rgba(201,24,74,.4) !important;
}

/* ── primary button ── */
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #FF6B6B, #C9184A) !important;
    color: white !important; border: none !important;
    border-radius: 14px !important; padding: 14px 28px !important;
    font-size: 1rem !important; font-weight: 700 !important;
    box-shadow: 0 4px 16px rgba(201,24,74,.3) !important;
}

/* ── dataframe ── */
.stDataFrame { border-radius: 14px !important; overflow: hidden !important; }

/* ── success box ── */
.stSuccess { border-radius: 12px !important; }

/* ── footer ── */
.footer {
    text-align: center; color: #bbb; font-size: .8rem;
    padding: 32px 0 16px;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────
CATEGORIES = {
    "Bao Cafe":         {"icon": "☕", "color": "#C2185B"},
    "อาหารพร้อมทาน":   {"icon": "🍱", "color": "#F57C00"},
    "เครื่องดื่ม":      {"icon": "🥤", "color": "#1976D2"},
    "ขนมขบเคี้ยว":     {"icon": "🍿", "color": "#7B1FA2"},
    "ของใช้ส่วนตัว":   {"icon": "🧴", "color": "#00897B"},
    "ของใช้ในบ้าน":    {"icon": "🏠", "color": "#558B2F"},
    "สินค้าเบ็ดเตล็ด": {"icon": "🔋", "color": "#6D4C41"},
    "บริการและอื่นๆ":  {"icon": "📱", "color": "#455A64"},
}

SYSTEM_PROMPT = """คุณคือระบบจำแนกประเภทสินค้าร้านสะดวกซื้อ CJ Express
จำแนกชื่อสินค้าแต่ละรายการออกเป็นหมวดหมู่ต่อไปนี้เท่านั้น:

1. Bao Cafe        → เครื่องดื่มกาแฟ ชา ลาเต้ เอสเปรสโซ่ อเมริกาโน่ มัทฉะ ชาเขียว ที่ขายในร้าน Bao Cafe (ชื่อสินค้ามักขึ้นต้นด้วย Bao หรือ Bac)
2. อาหารพร้อมทาน  → อาหารแช่แข็ง แซนด์วิช เบเกอรี่ ไส้กรอก ปัง ขนมปัง อาหารสำเร็จรูป พุดดิ้ง ข้าวเกรียบ ผงชูรส ไข่ อาหารคาว
3. เครื่องดื่ม     → น้ำดื่ม น้ำอัดลม นม ชาทั่วไป กาแฟทั่วไป เบียร์ สุรา น้ำหวาน โซดา เครื่องดื่มชูกำลัง คอลลาเจนดื่ม โพรไบโอติกส์
4. ขนมขบเคี้ยว    → มันฝรั่งทอด ลูกอม ช็อกโกแลต ไอศกรีม ขนมกรอบ ถั่ว เมล็ดทานตะวัน อัลมอนด์ แครกเกอร์ ข้าวโพดอบ
5. ของใช้ส่วนตัว  → ยาสีฟัน แชมพู สบู่ ครีมทาผิว ผ้าอนามัย โฟมล้างหน้า รองพื้น เครื่องสำอาง แปรงสีฟัน ยา วิตามิน เจล
6. ของใช้ในบ้าน   → กระดาษทิชชู ผงซักฟอก น้ำยาล้างจาน ถุงขยะ น้ำยาทำความสะอาด ปรับผ้านุ่ม
7. สินค้าเบ็ดเตล็ด → ถ่านไฟฉาย ไฟแช็ก ยากันยุง อุปกรณ์เครื่องเขียน เทปลบคำผิด ยางลบ
8. บริการและอื่นๆ  → เติมเงินมือถือ จ่ายบิล ซิมการ์ด บัตรเติมเงิน ส่วนลด แก้วลูกค้า

ตอบเป็น JSON เท่านั้น ไม่มี markdown: {"ชื่อสินค้า1":"ประเภท1",...}
ถ้าไม่แน่ใจใส่ "สินค้าเบ็ดเตล็ด"
"""

# ─────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────
def get_api_key():
    import os
    # วิธีที่ 1: st.secrets direct key
    try:
        key = st.secrets["ANTHROPIC_API_KEY"]
        if key and str(key).strip(): return str(key).strip()
    except: pass
    # วิธีที่ 2: nested section
    try:
        key = st.secrets["secrets"]["ANTHROPIC_API_KEY"]
        if key and str(key).strip(): return str(key).strip()
    except: pass
    # วิธีที่ 3: environment variable
    key = os.environ.get("ANTHROPIC_API_KEY", "")
    if key and key.strip(): return key.strip()
    return ""

def claude_classify(products, api_key):
    payload = json.dumps({
        "model": "claude-sonnet-4-6",
        "max_tokens": 2000,
        "system": SYSTEM_PROMPT,
        "messages": [{"role":"user","content":"จำแนกสินค้า:\n"+"\n".join(f"- {p}" for p in products)}],
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
        status.markdown(f"🤖 วิเคราะห์ **{min((i+1)*batch_size,len(products))}/{len(products)}** รายการ...")
        try:    result.update(claude_classify(batch, api_key))
        except Exception as e:
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

def sec(title, icon=""):
    st.markdown(f'<div class="sec-title"><div class="dot"></div>{icon} {title}</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────
for k, v in [("df",None),("cat_map",{}),("analyzed",False)]:
    if k not in st.session_state: st.session_state[k] = v

# ─────────────────────────────────────────
# HERO
# ─────────────────────────────────────────
st.markdown("""
<div class="hero">
  <div class="hero-mascot">🛒</div>
  <div class="hero-title">CJ Smart Scan</div>
  <div class="hero-sub">📊 ระบบวิเคราะห์และจัดหมวดหมู่สินค้าอัตโนมัติด้วย AI</div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# API KEY CHECK
# ─────────────────────────────────────────
api_key = get_api_key()
if not api_key:
    st.error("⚠️ ไม่พบ **ANTHROPIC_API_KEY**")
    # Debug: แสดงสิ่งที่อ่านได้จาก secrets
    with st.expander("🔍 Debug Info"):
        try:
            keys = list(st.secrets.keys())
            st.write("Keys ใน secrets:", keys)
            for k in keys:
                val = str(st.secrets[k])
                st.write(f"{k} = {val[:10]}...{val[-4:]} (length: {len(val)})")
        except Exception as e:
            st.write("อ่าน secrets ไม่ได้:", e)
    st.info("วิธีแก้: ไปที่ App Settings → Secrets แล้วใส่\n```\nANTHROPIC_API_KEY = \"sk-ant-...\"\n```\nแล้วกด Save และรอ 1 นาที")
    st.stop()

# ─────────────────────────────────────────
# UPLOAD + RESULT in two-column layout
# ─────────────────────────────────────────
left, right = st.columns([1, 1.6], gap="large")

with left:
    sec("อัปโหลดไฟล์ Excel", "📂")
    uploaded = st.file_uploader(
        "ลากไฟล์มาวางหรือกดเลือกไฟล์",
        type=["xlsx"],
        label_visibility="collapsed",
    )
    if uploaded:
        try:
            df_raw = pd.read_excel(uploaded, sheet_name="ใบเสร็จ")
            if st.session_state.df is None or st.session_state.df.shape != df_raw.shape:
                st.session_state.df = df_raw
                st.session_state.analyzed = False
                st.session_state.cat_map = {}
            n_items = df_raw["ชื่อสินค้า"].notna().sum()
            branches = df_raw["รหัสสาขา"].dropna().nunique()
            st.success(f"✅ โหลดสำเร็จ! พบ **{n_items}** รายการสินค้า จาก **{branches}** สาขา")
        except Exception as e:
            st.error(f"❌ {e}"); st.stop()

    if st.session_state.df is not None and not st.session_state.analyzed:
        if st.button("🤖  วิเคราะห์ประเภทสินค้า", type="primary", use_container_width=True):
            products = st.session_state.df["ชื่อสินค้า"].dropna().unique().tolist()
            st.session_state.cat_map = classify_all(products, api_key)
            st.session_state.analyzed = True
            st.rerun()

    # ── Category legend ──
    if st.session_state.analyzed:
        sec("หมวดหมู่สินค้า", "🏷️")
        for name, meta in CATEGORIES.items():
            st.markdown(
                f'<div style="display:flex;align-items:center;gap:10px;padding:6px 0;">'
                f'<span style="font-size:1.4rem">{meta["icon"]}</span>'
                f'<span class="cat-pill" style="background:{meta["color"]}">{name}</span>'
                f'</div>',
                unsafe_allow_html=True
            )

with right:
    if st.session_state.analyzed and st.session_state.df is not None:
        df = st.session_state.df.copy()
        cat_map = st.session_state.cat_map
        df["ประเภทสินค้า"] = df["ชื่อสินค้า"].map(cat_map)
        items = df[df["ชื่อสินค้า"].notna()].copy()

        # ── Download button (top) ──
        summary_df = items.groupby("ประเภทสินค้า").agg(
            จำนวนรายการ=("ชื่อสินค้า","count"), ยอดรวม=("ยอดรวมสินค้า","sum")
        ).reset_index().sort_values("จำนวนรายการ", ascending=False)

        branch_df = items.groupby(["รหัสสาขา","ประเภทสินค้า"]).agg(
            จำนวนรายการ=("ชื่อสินค้า","count"), ยอดรวม=("ยอดรวมสินค้า","sum")
        ).reset_index().sort_values(["รหัสสาขา","จำนวนรายการ"], ascending=[True,False])

        map_df = pd.DataFrame(sorted(cat_map.items(), key=lambda x:x[1]),
                              columns=["ชื่อสินค้า","ประเภทสินค้า"])

        excel_bytes = build_excel(df, summary_df, branch_df, map_df)

        st.download_button(
            "⬇️  ดาวน์โหลด Excel สรุปผลทั้งหมด",
            data=excel_bytes,
            file_name="CJ_product_classified.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.markdown("<br>", unsafe_allow_html=True)

        # ── Metric cards ──
        total_sales = items["ยอดรวมสินค้า"].sum()
        m1,m2,m3,m4 = st.columns(4)
        for col, icon, num, lbl in [
            (m1,"🏪", items["รหัสสาขา"].nunique(), "สาขา"),
            (m2,"📦", len(items), "รายการ"),
            (m3,"🏷️", items["ประเภทสินค้า"].nunique(), "หมวดหมู่"),
            (m4,"💰", f"฿{total_sales:,.0f}", "ยอดรวม"),
        ]:
            col.markdown(f"""
            <div class="mcard">
              <div class="icon">{icon}</div>
              <div class="num">{num}</div>
              <div class="lbl">{lbl}</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Category summary bars ──
        sec("สัดส่วนตามหมวดหมู่", "📊")
        max_cnt = summary_df["จำนวนรายการ"].max()
        for _, row in summary_df.iterrows():
            cat   = row["ประเภทสินค้า"]
            cnt   = int(row["จำนวนรายการ"])
            total = row["ยอดรวม"]
            meta  = CATEGORIES.get(cat, {"icon":"📦","color":"#9E9E9E"})
            pct   = cnt / max_cnt * 100
            st.markdown(f"""
            <div class="cat-card">
              <div class="cat-icon">{meta["icon"]}</div>
              <div style="flex:1">
                <div class="cat-name">{cat}</div>
                <div class="cat-count">฿{total:,.0f}</div>
                <div class="bar-wrap" style="margin-top:6px">
                  <div class="bar-fill" style="width:{pct}%;background:{meta['color']}"></div>
                </div>
              </div>
              <div class="cat-num">{cnt}</div>
            </div>""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# DETAIL TABLE (full width)
# ─────────────────────────────────────────
if st.session_state.analyzed and st.session_state.df is not None:
    df = st.session_state.df.copy()
    df["ประเภทสินค้า"] = df["ชื่อสินค้า"].map(st.session_state.cat_map)
    items = df[df["ชื่อสินค้า"].notna()].copy()

    st.markdown("<br>", unsafe_allow_html=True)
    sec("รายการสินค้าทั้งหมด", "📋")

    fa, fb, fc = st.columns([1,1,1])
    branch_ids = sorted(items["รหัสสาขา"].dropna().unique().astype(int).tolist())
    with fa:
        sel_branch = st.multiselect("🏪 สาขา", branch_ids, default=branch_ids,
                                    format_func=lambda x: f"สาขา {x}")
    with fb:
        all_cats = list(CATEGORIES.keys())
        sel_cat = st.multiselect("🏷️ ประเภท", all_cats, default=all_cats)
    with fc:
        search = st.text_input("🔍 ค้นหาชื่อสินค้า", placeholder="พิมพ์ชื่อสินค้า...")

    filtered = items[
        items["รหัสสาขา"].isin(sel_branch) &
        items["ประเภทสินค้า"].isin(sel_cat)
    ].copy()
    if search:
        filtered = filtered[filtered["ชื่อสินค้า"].str.contains(search, na=False, case=False)]

    show_cols = ["รหัสสาขา","วันที่","ชื่อสินค้า","ประเภทสินค้า","จำนวน","ราคาต่อหน่วย","ยอดรวมสินค้า"]
    st.dataframe(
        filtered[show_cols].style.format({
            "จำนวน":"{:.0f}", "ราคาต่อหน่วย":"฿{:,.2f}", "ยอดรวมสินค้า":"฿{:,.2f}"
        }).apply(lambda col: [
            f"background-color:{CATEGORIES.get(v,{}).get('color','#9E9E9E')}22; color:{CATEGORIES.get(v,{}).get('color','#555')};font-weight:600"
            if col.name == "ประเภทสินค้า" else "" for v in col
        ], axis=0),
        use_container_width=True, hide_index=True, height=420,
    )
    st.caption(f"แสดง **{len(filtered):,}** จาก **{len(items):,}** รายการ")

    # ── per-branch tabs ──
    st.markdown("<br>", unsafe_allow_html=True)
    sec("สรุปรายสาขา", "🏪")
    branch_df2 = items.groupby(["รหัสสาขา","ประเภทสินค้า"]).agg(
        จำนวนรายการ=("ชื่อสินค้า","count"), ยอดรวม=("ยอดรวมสินค้า","sum")
    ).reset_index()

    tabs = st.tabs([f"🏪 สาขา {b}" for b in branch_ids])
    for tab, b in zip(tabs, branch_ids):
        with tab:
            bd = branch_df2[branch_df2["รหัสสาขา"]==b].copy()
            c1, c2 = st.columns([1.2,1])
            with c1:
                st.bar_chart(bd.set_index("ประเภทสินค้า")["จำนวนรายการ"],
                             color="#FF4D6D", height=260)
            with c2:
                st.dataframe(
                    bd[["ประเภทสินค้า","จำนวนรายการ","ยอดรวม"]]
                    .style.format({"ยอดรวม":"฿{:,.2f}"}),
                    use_container_width=True, hide_index=True, height=260
                )

# ─────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────
st.markdown('<div class="footer">🛒 CJ Smart Scan · Powered by Claude AI · Made with ❤️</div>',
            unsafe_allow_html=True)
