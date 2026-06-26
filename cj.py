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
*,[class*="css"]{font-family:'Sarabun',sans-serif!important}
#MainMenu,header,footer{visibility:hidden}
.stApp{background:linear-gradient(160deg,#FFF8F0 0%,#FFF0F5 50%,#F0F4FF 100%)}
.hero{background:linear-gradient(135deg,#FF6B6B 0%,#FF4D6D 40%,#C9184A 100%);
  border-radius:24px;padding:32px 40px 28px;margin-bottom:24px;position:relative;
  overflow:hidden;box-shadow:0 8px 32px rgba(201,24,74,.25)}
.hero-title{font-family:'Fredoka One','Sarabun',sans-serif!important;
  font-size:2.4rem;font-weight:900;color:white;margin:0;text-shadow:0 2px 8px rgba(0,0,0,.2)}
.hero-sub{color:rgba(255,255,255,.88);font-size:1rem;margin:8px 0 0}
.stDownloadButton>button, .stButton>button[kind="primary"]{
  background:linear-gradient(135deg,#FF6B6B,#C9184A)!important;
  color:white!important;border:none!important;border-radius:12px!important;
  padding:12px 24px!important;font-weight:700!important}
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

# ========================= SMART RULE KW =========================
RULE_KW = {
    "Fresh Food": {
        "keywords": ["bao","bac","แซนด์วิช","ไส้กรอก","ขนมปัง","เบเกอรี","sandwich","sausage","ข้าวกล่อง","สลัด","salad","นม pasteur","โยเกิร์ต","ไก่","หมู","เนื้อ","ทูน่า","ปลา","กริล","อบ","ปิ้ง","โรตี","ครัวซอง","โดนัท","เค้ก","พาย","fresh","bakery","counter"],
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

# ========================= SYSTEM PROMPT =========================
SYSTEM_PROMPT = """คุณคือผู้เชี่ยวชาญการจำแนกสินค้า CJ Express ตาม PMA 5 กลุ่มอย่างเคร่งครัด

1. Fresh Food → อาหารสด เบเกอรี่สด แซนด์วิช ไส้กรอก ข้าวกล่อง สลัด นมพาสเจอร์ไรส์
2. Non Food → ของใช้ส่วนตัว ของใช้ในบ้าน บุหรี่ เครื่องเขียน ทิชชู ผ้าอนามัย
3. Packaged Beverage → น้ำดื่ม น้ำอัดลม เบียร์ สุรา ไอศกรีม นม UHT
4. Processed Food → อาหารแปรรูป บะหมี่ ขนมขบเคี้ยว ช็อกโกแลต อาหารกระป๋อง
5. Special Business → ยา วิตามิน บริการเติมเงิน จ่ายบิล Bellinee Kudsan

ตอบเป็น JSON เท่านั้น: {"ชื่อสินค้า": "ชื่อกลุ่ม", ...}
ห้ามตอบกลุ่มอื่นนอกเหนือจาก 5 กลุ่มนี้"""

# ========================= HELPERS =========================
def get_api_key():
    import os
    try:
        return st.secrets["GEMINI_API_KEY"].strip()
    except:
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
        raw = re.sub(r"```json|```", "", data["candidates"][0]["content"]["parts"][0]["text"]).strip()
        result = json.loads(raw)
        valid = set(CATEGORIES.keys())
        return {k: (v if v in valid else "Processed Food") for k, v in result.items()}
    except Exception as e:
        st.warning(f"Gemini Error: {str(e)[:100]}")
        return {}

def post_validation(cat_map):
    validated = {}
    for product, category in cat_map.items():
        p = str(product).lower()
        if any(x in p for x in ["ไอศกรีม", "ice cream"]):
            validated[product] = "Packaged Beverage"
        elif any(x in p for x in ["ขนมปัง", "เบเกอรี", "bakery", "โดนัท", "เค้ก", "โรตี"]):
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
            bar.progress((i+1)/len(batches))
        
        bar.empty()
        status.empty()
    
    result = post_validation(result)
    return result

# ========================= EXCEL BUILDER (คงเดิม) =========================
# ... (ส่วน build_excel, load_excel, make_summary ฯลฯ ยังใช้โค้ดเดิมของคุณ) ...
# เพื่อความกระชับ ฉันข้ามส่วนยาวๆ ไว้ก่อน คุณสามารถ copy ส่วนนี้จากโค้ดเดิมของคุณมาใส่ได้เลย

# ========================= SESSION STATE =========================
for k in ["df", "cat_map", "analyzed", "_fid", "df_prev", "cat_map_prev", "analyzed_prev", "_fid_prev"]:
    if k not in st.session_state:
        st.session_state[k] = None if "df" in k else False if "analyzed" in k else ""

api_key = get_api_key()

# ========================= UI =========================
st.markdown("""
<div class="hero">
  <div class="hero-title">CJ Smart Scan</div>
  <div class="hero-sub">📊 วิเคราะห์ยอดลูกค้าสะสม + จำแนก PMA 5 กลุ่ม (Smart Hybrid)</div>
</div>
""", unsafe_allow_html=True)

if not api_key:
    st.error("⚠️ ไม่พบ GEMINI_API_KEY")
    st.stop()

tab1, tab2 = st.tabs(["📦 วิเคราะห์ไฟล์", "📊 เปรียบเทียบเดือน"])

with tab1:
    left, right = st.columns([1, 1.6])
    with left:
        uploaded = st.file_uploader("อัปโหลดไฟล์ Excel", type=["xlsx"], key="up1")
        if uploaded:
            fid = f"{uploaded.name}_{uploaded.size}"
            if st.session_state._fid != fid:
                df_raw = pd.read_excel(uploaded, sheet_name="ใบเสร็จ")
                st.session_state.df = df_raw
                st.session_state.analyzed = False
                st.session_state.cat_map = {}
                st.session_state._fid = fid

        if st.session_state.df is not None and not st.session_state.analyzed:
            if st.button("⚡ จำแนกสินค้า (Smart Hybrid)", type="primary", use_container_width=True):
                prods = st.session_state.df["ชื่อสินค้า"].dropna().unique().tolist()
                with st.spinner("กำลังจำแนกด้วย Smart Hybrid Logic..."):
                    st.session_state.cat_map = smart_hybrid_classify(prods, api_key, st.session_state.cat_map)
                    st.session_state.analyzed = True
                st.success("✅ จำแนกสินค้าเสร็จสิ้น!")
                st.rerun()

        sec("5 กลุ่ม PMA","🏷️")
        for name,meta in CATEGORIES.items():
            st.markdown(
                f'<div style="display:flex;align-items:center;gap:10px;padding:5px 0">'
                f'<span style="font-size:1.3rem">{meta["icon"]}</span>'
                f'<span style="background:#{meta["hex"]}18;color:#{meta["hex"]};padding:4px 14px;'
                f'border-radius:99px;font-size:.82rem;font-weight:700;border:1px solid #{meta["hex"]}44">'
                f'{name}</span>'
                f'<span style="font-size:.72rem;color:#aaa">{len(meta["sub"])} sub-groups</span>'
                f'</div>', unsafe_allow_html=True)

    with right:
        if st.session_state.analyzed and st.session_state.df is not None:
            df=st.session_state.df.copy()
            df["ประเภทสินค้า"]=df["ชื่อสินค้า"].map(st.session_state.cat_map)
            items=df[df["ชื่อสินค้า"].notna()].copy()

            # ยอดลูกค้าสะสม
            mc_df, branch_tot_df = make_branch_customer_summary(df)
            summary_df=make_summary(items)
            map_df=pd.DataFrame(sorted(st.session_state.cat_map.items(),key=lambda x:x[1]),
                                columns=["ชื่อสินค้า","ประเภทสินค้า"])

            # prev
            has_prev=st.session_state.df_prev is not None and st.session_state.analyzed_prev
            sp=mc_prev=bp=None
            if has_prev:
                dfp=st.session_state.df_prev.copy()
                dfp["ประเภทสินค้า"]=dfp["ชื่อสินค้า"].map(st.session_state.cat_map_prev)
                ip=dfp[dfp["ชื่อสินค้า"].notna()].copy()
                sp=make_summary(ip)
                mc_prev,bp=make_branch_customer_summary(dfp)

            excel_bytes=build_excel(df,summary_df,items,mc_df,branch_tot_df,map_df,
                                    df_prev=st.session_state.df_prev if has_prev else None,
                                    summary_prev=sp, mc_prev=mc_prev, branch_prev=bp)
            lbl="⬇️ ดาวน์โหลด Excel (พร้อมเปรียบเทียบ)" if has_prev else "⬇️ ดาวน์โหลด Excel"
            st.download_button(lbl,data=excel_bytes,file_name="CJ_SmartScan.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)
            st.markdown("<br>",unsafe_allow_html=True)

            # metrics
            total_cust=int(branch_tot_df["ยอดลูกค้าสะสมรวม"].sum())
            m1,m2,m3,m4=st.columns(4)
            for col,icon,num,lbl2 in [
                (m1,"🏪",b,"สาขา"),
                (m2,"👥",f"{total_cust:,}","ยอดลูกค้าสะสม"),
                (m3,"📦",f"{len(items):,}","รายการสินค้า"),
                (m4,"🏷️",int(items["ประเภทสินค้า"].nunique()),"กลุ่ม PMA"),
            ]:
                col.markdown(f'<div class="mcard"><div class="icon">{icon}</div>'
                             f'<div class="num">{num}</div><div class="lbl">{lbl2}</div></div>',
                             unsafe_allow_html=True)

    # ── Full-width sections ──
    if st.session_state.analyzed and st.session_state.df is not None:
        df2=st.session_state.df.copy()
        df2["ประเภทสินค้า"]=df2["ชื่อสินค้า"].map(st.session_state.cat_map)
        items2=df2[df2["ชื่อสินค้า"].notna()].copy()
        mc_df2,branch_tot_df2=make_branch_customer_summary(df2)

        # AI button
        prods=st.session_state.df["ชื่อสินค้า"].dropna().unique().tolist()
        need_ai=[p for p in prods if not rule_classify(p)]
        st.markdown("<br>",unsafe_allow_html=True)
        if st.session_state.get("ai_status")=="done":
            st.success("🎉 AI จำแนกครบทุกรายการแล้ว!")
            st.session_state.ai_status=""
        if need_ai:
            ci,cb=st.columns([2,1])
            ci.info(f"มี **{len(need_ai)}** รายการที่ใช้ rule-based — กด AI เพื่อเพิ่มความแม่นยำ")
            if cb.button("🤖 จำแนกเพิ่มด้วย Gemini AI",use_container_width=True,key="btn_ai"):
                with st.spinner("🤖 กำลังวิเคราะห์..."):
                    new_map=classify_with_ai(prods,api_key,st.session_state.cat_map)
                st.session_state.cat_map=new_map
                st.session_state.ai_status="done"
                st.rerun()
        else:
            st.success("✅ จำแนกครบทุกรายการ 🎉")

        # ── ยอดลูกค้าสะสมรายสาขา (MAIN FOCUS) ──
        st.markdown("<br>",unsafe_allow_html=True)
        sec("👥 ยอดลูกค้าสะสมรายสาขา","🏆")
        st.caption("ยอดลูกค้าสะสม = ผลรวม max transaction no. ของแต่ละเครื่องคิดเงิน ต่อสาขา")

        top_n=min(10,len(branch_tot_df2))
        max_cust=int(branch_tot_df2["ยอดลูกค้าสะสมรวม"].max()) or 1
        col_l,col_r=st.columns([1.4,1],gap="large")
        with col_l:
            for i,(_,row) in enumerate(branch_tot_df2.head(top_n).iterrows()):
                bid=int(row["รหัสสาขา"]); cv=int(row["ยอดลูกค้าสะสมรวม"])
                pct=cv/max_cust*100
                medal=RANK_MEDAL[i] if i<len(RANK_MEDAL) else str(i+1)
                rbg=RANK_BG[i] if i<len(RANK_BG) else "AAAAAA"
                # หาเครื่องของสาขานี้
                mach_rows=mc_df2[mc_df2["รหัสสาขา"]==row["รหัสสาขา"]]
                mach_txt=" | ".join([f"{r['เครื่อง']}:{int(r['ยอดลูกค้าสะสม']):,}" for _,r in mach_rows.iterrows()])
                st.markdown(
                    f'<div class="branch-card">'
                    f'<div style="display:flex;align-items:center;gap:14px">'
                    f'<div style="background:#{rbg};color:{"#333" if rbg in ["FFD700","C0C0C0","CD7F32"] else "#fff"};'
                    f'width:40px;height:40px;border-radius:50%;display:flex;align-items:center;'
                    f'justify-content:center;font-size:1.1rem;font-weight:900;flex-shrink:0">{medal}</div>'
                    f'<div style="flex:1">'
                    f'<div style="font-weight:800;font-size:1rem">🏪 สาขา {bid}</div>'
                    f'<div style="font-size:.75rem;color:#aaa;margin:2px 0 6px">{mach_txt}</div>'
                    f'<div style="background:#f0f0f0;border-radius:99px;height:8px;overflow:hidden">'
                    f'<div style="width:{pct}%;height:100%;background:#C9184A;border-radius:99px"></div></div>'
                    f'</div>'
                    f'<div style="font-weight:900;font-size:1.4rem;color:#C9184A;min-width:60px;text-align:right">'
                    f'{cv:,}</div>'
                    f'</div></div>', unsafe_allow_html=True)
        with col_r:
            st.dataframe(
                branch_tot_df2.rename(columns={"ยอดลูกค้าสะสมรวม":"ยอดลูกค้าสะสม"})
                .style.format({"ยอดลูกค้าสะสม":"{:,.0f}"}),
                use_container_width=True, hide_index=True, height=420)

        # ── ยอดลูกค้าแยกเครื่อง ──
        st.markdown("<br>",unsafe_allow_html=True)
        sec("🖨️ ยอดลูกค้าสะสมแยกตามเครื่องคิดเงิน","📊")
        bids_all=sorted(mc_df2["รหัสสาขา"].dropna().unique().astype(int).tolist())
        if bids_all:
            tabs_b=st.tabs([f"สาขา {b}" for b in bids_all[:10]])
            for tab,bid in zip(tabs_b,bids_all[:10]):
                with tab:
                    bd=mc_df2[mc_df2["รหัสสาขา"]==bid].sort_values("เครื่อง")
                    c1,c2=st.columns([1,1])
                    with c1:
                        st.bar_chart(bd.set_index("เครื่อง")["ยอดลูกค้าสะสม"],
                                     color="#C9184A",height=280)
                    with c2:
                        st.dataframe(bd.style.format({"ยอดลูกค้าสะสม":"{:,.0f}"}),
                                     use_container_width=True,hide_index=True,height=280)
                        total_b=int(bd["ยอดลูกค้าสะสม"].sum())
                        st.metric("ยอดลูกค้าสะสมรวมสาขานี้",f"{total_b:,}")

        # ── สัดส่วนสินค้าตาม PMA ──
        st.markdown("<br>",unsafe_allow_html=True)
        sec("📊 สัดส่วนสินค้าตาม PMA 5 กลุ่ม","🏷️")
        summary_df2=make_summary(items2)
        cb1,cb2=st.columns([1,1],gap="large")
        with cb1:
            max_cnt=int(summary_df2["จำนวนรายการ"].max()) or 1
            for _,row in summary_df2.iterrows():
                cat=row["ประเภทสินค้า"]; cnt=int(row["จำนวนรายการ"])
                meta=CATEGORIES.get(cat,{"icon":"📦","color":"#9E9E9E","hex":"9E9E9E"})
                pct=cnt/max_cnt*100
                st.markdown(
                    f'<div class="cat-card">'
                    f'<div class="cat-icon">{meta["icon"]}</div>'
                    f'<div style="flex:1"><div class="cat-name">{cat}</div>'
                    f'<div class="bar-wrap" style="margin-top:5px">'
                    f'<div class="bar-fill" style="width:{pct}%;background:#{meta["hex"]}"></div>'
                    f'</div></div><div class="cat-num">{cnt:,}</div></div>',
                    unsafe_allow_html=True)
        with cb2:
            st.dataframe(summary_df2.style.format({"จำนวนรายการ":"{:,.0f}","ยอดรวม":"฿{:,.2f}"}),
                         use_container_width=True,hide_index=True,height=320)

        # ── ตารางรายการ ──
        st.markdown("<br>",unsafe_allow_html=True)
        sec("📋 รายการสินค้าทั้งหมด")
        fa,fb,fc=st.columns(3)
        bids2=sorted(items2["รหัสสาขา"].dropna().unique().astype(int).tolist())
        with fa: sel_b=st.multiselect("🏪 สาขา",bids2,default=bids2,key="sb1",format_func=lambda x:f"สาขา {x}")
        with fb: sel_c=st.multiselect("🏷️ กลุ่ม PMA",list(CATEGORIES.keys()),default=list(CATEGORIES.keys()),key="sc1")
        with fc: srch=st.text_input("🔍 ค้นหา",placeholder="ชื่อสินค้า...",key="sr1")
        filt=items2[items2["รหัสสาขา"].isin(sel_b)&items2["ประเภทสินค้า"].isin(sel_c)].copy()
        if srch: filt=filt[filt["ชื่อสินค้า"].str.contains(srch,na=False,case=False)]
        showcols=["รหัสสาขา","วันที่","ชื่อสินค้า","ประเภทสินค้า","จำนวน","ราคาต่อหน่วย","ยอดรวมสินค้า"]
        showcols=[c for c in showcols if c in filt.columns]
        def hl(v):
            m=CATEGORIES.get(v,{"color":"#9E9E9E","hex":"9E9E9E"})
            return f"background:#{m['hex']}15;color:#{m['hex']};font-weight:700"
        st.dataframe(filt[showcols].style.map(hl,subset=["ประเภทสินค้า"]),
                     use_container_width=True,hide_index=True,height=420)
        st.caption(f"แสดง **{len(filt):,}** จาก **{len(items2):,}** รายการ")

# ══════════════════════
# TAB 2: เปรียบเทียบ
# ══════════════════════
with tab2:
    st.markdown("### 📊 เปรียบเทียบยอดลูกค้าสะสมกับเดือนก่อน")
    col_up,col_info=st.columns([1,1.6],gap="large")
    with col_up:
        sec("อัปโหลดไฟล์เดือนก่อน","📂")
        up2=st.file_uploader("เลือกไฟล์ Excel เดือนก่อน",type=["xlsx"],
                              key="up2",label_visibility="collapsed")
        if up2 is not None:
            fid2=f"{up2.name}_{up2.size}"
            if st.session_state._fid_prev!=fid2:
                try:
                    df_p=load_excel(up2)
                    st.session_state.df_prev=df_p
                    st.session_state.analyzed_prev=False
                    st.session_state.cat_map_prev={}
                    st.session_state._fid_prev=fid2
                except Exception as e:
                    st.error(f"❌ {e}"); st.stop()

        if st.session_state.df_prev is not None:
            n2=int(st.session_state.df_prev["ชื่อสินค้า"].notna().sum())
            b2=int(st.session_state.df_prev["รหัสสาขา"].dropna().nunique())
            st.success(f"✅ พบ **{n2:,}** รายการ จาก **{b2}** สาขา")
        if st.session_state.df_prev is not None and not st.session_state.analyzed_prev:
            if st.button("⚡ จำแนกไฟล์เดือนก่อน",type="primary",use_container_width=True,key="btn_prev"):
                prods2=st.session_state.df_prev["ชื่อสินค้า"].dropna().unique().tolist()
                st.session_state.cat_map_prev=classify_rule_only(prods2)
                st.session_state.analyzed_prev=True
                st.rerun()
        if not st.session_state.analyzed:
            st.info("💡 จำแนกไฟล์เดือนปัจจุบันในแท็บแรกก่อน")

    with col_info:
        if st.session_state.analyzed and st.session_state.analyzed_prev:
            df_cur=st.session_state.df.copy()
            df_cur["ประเภทสินค้า"]=df_cur["ชื่อสินค้า"].map(st.session_state.cat_map)
            items_cur=df_cur[df_cur["ชื่อสินค้า"].notna()].copy()
            mc_cur,bt_cur=make_branch_customer_summary(df_cur)

            df_prv=st.session_state.df_prev.copy()
            df_prv["ประเภทสินค้า"]=df_prv["ชื่อสินค้า"].map(st.session_state.cat_map_prev)
            items_prv=df_prv[df_prv["ชื่อสินค้า"].notna()].copy()
            mc_prv,bt_prv=make_branch_customer_summary(df_prv)

            lbl_cur=items_cur["เดือน"].iloc[0] if len(items_cur) else "ปัจจุบัน"
            lbl_prv=items_prv["เดือน"].iloc[0] if len(items_prv) else "ก่อนหน้า"

            tot_cur=int(bt_cur["ยอดลูกค้าสะสมรวม"].sum())
            tot_prv=int(bt_prv["ยอดลูกค้าสะสมรวม"].sum())
            chg=tot_cur-tot_prv
            chg_pct=(chg/tot_prv*100) if tot_prv>0 else 0
            arrow="📈" if chg>=0 else "📉"
            color="#2E7D32" if chg>=0 else "#C62828"

            m1,m2,m3=st.columns(3)
            m1.markdown(f'<div class="mcard"><div class="icon">📅</div>'
                        f'<div class="num" style="font-size:1rem">{lbl_prv}</div>'
                        f'<div class="lbl">เดือนก่อน: {tot_prv:,}</div></div>',unsafe_allow_html=True)
            m2.markdown(f'<div class="mcard"><div class="icon">📅</div>'
                        f'<div class="num" style="font-size:1rem">{lbl_cur}</div>'
                        f'<div class="lbl">ปัจจุบัน: {tot_cur:,}</div></div>',unsafe_allow_html=True)
            m3.markdown(f'<div class="mcard"><div class="icon">{arrow}</div>'
                        f'<div class="num" style="color:{color}">{chg_pct:+.1f}%</div>'
                        f'<div class="lbl">ยอดลูกค้าสะสมเปลี่ยน</div></div>',unsafe_allow_html=True)

            st.markdown("<br>",unsafe_allow_html=True)
            sec("🏪 เปรียบเทียบยอดลูกค้าสะสมรายสาขา")
            mrg=bt_cur.rename(columns={"ยอดลูกค้าสะสมรวม":"ปัจจุบัน"}).merge(
                bt_prv.rename(columns={"ยอดลูกค้าสะสมรวม":"เดือนก่อน"}),
                on="รหัสสาขา",how="outer").fillna(0)
            mrg["เปลี่ยนแปลง"]=mrg["ปัจจุบัน"]-mrg["เดือนก่อน"]
            mrg["chg_pct"]=mrg.apply(
                lambda r: (r["เปลี่ยนแปลง"]/r["เดือนก่อน"]*100) if r["เดือนก่อน"]>0 else 100, axis=1)
            mrg=mrg.sort_values("ปัจจุบัน",ascending=False)
            for _,row in mrg.head(10).iterrows():
                bid=int(row["รหัสสาขา"]); up=row["เปลี่ยนแปลง"]>=0
                c2=("#2E7D32" if up else "#C62828")
                bg2=("E8F5E9" if up else "FFEBEE")
                pct_str=f'{row["chg_pct"]:+.1f}%'
                st.markdown(
                    f'<div class="compare-card">'
                    f'<div style="display:flex;justify-content:space-between;align-items:center">'
                    f'<div><div style="font-weight:700;font-size:1rem">🏪 สาขา {bid}</div>'
                    f'<div style="font-size:.8rem;color:#888">'
                    f'{lbl_prv}: {int(row["เดือนก่อน"]):,} → {lbl_cur}: {int(row["ปัจจุบัน"]):,} ลูกค้า</div></div>'
                    f'<div style="text-align:right">'
                    f'<div style="font-size:1.6rem">{"📈" if up else "📉"}</div>'
                    f'<span style="background:#{bg2};color:{c2};padding:3px 14px;border-radius:99px;'
                    f'font-weight:700;font-size:.9rem">{pct_str}</span>'
                    f'</div></div></div>', unsafe_allow_html=True)
        else:
            st.info("📌 อัปโหลดและจำแนกไฟล์ทั้งสองเดือนก่อนเพื่อดูผลเปรียบเทียบ")

st.markdown('<div class="footer">🛒 CJ Smart Scan · PMA 5 Groups · ยอดลูกค้าสะสม · Powered by Gemini AI</div>',
            unsafe_allow_html=True)
