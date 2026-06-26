import streamlit as st
import pandas as pd
import json, re, io, urllib.request

# ══════════════════════════════════════
# CONFIG & SYSTEM PROMPT (ปรับปรุงตามไฟล์ PMA ของคุณ)
# ══════════════════════════════════════
st.set_page_config(page_title="CJ Smart Scan Pro", page_icon="🛒", layout="wide")

SYSTEM_PROMPT = """คุณคือระบบผู้ช่วยอัจฉริยะของ CJ Express หน้าที่ของคุณคือ:
1. Normalize: แก้ไขชื่อสินค้าที่ย่อให้เป็นชื่อเต็มที่อ่านเข้าใจง่าย
2. Classify: จัดกลุ่มสินค้าให้ตรงกับกลุ่ม PMA 5 กลุ่มหลักดังนี้:

- Fresh Food: (CHILLED BREAD, APPETIZER, FRUIT, COUNTER DRINK, FOOD PLACE, SANDWICH, BURGER, สลัด, MEAL BOX, SAUSAGE, PASTEURIZED MILK, BAKERY, READY TO COOK)
- Non Food: (CIGARETTE, BOOKS, ENTERTAINMENT, IT Devices, PERSONAL CARE, HOUSEWARE, STATIONERY, SANITARY, HOUSEHOLD, ELECTRONIC, HERBAL, สินค้าเทศกาล)
- Packaged Beverage: (BEER/ALCOHOL, LIQUOR, UHT MILK, NON-CARBONATED/CARBONATED SOFT DRINK, ICE CREAM, ICE, ENERGY/SPORT DRINK)
- Processed Food: (COOKING/CANNED FOODS, INSTANT FOODS, PACKAGED FOODS, CONFECTIONERY, SNACKS)
- Special Business: (HEALTH CARE, MEDICINE, SYNERGY PROJECT, 7-SERVICE, WELLNESS, BELLINEE, KUDSAN, SOCIAL WELFARE, HOME & LIVING)

ตอบกลับเป็น JSON เท่านั้นในรูปแบบ:
{"ชื่อสินค้าต้นทาง": {"full_name": "ชื่อที่แก้ไขแล้ว", "category": "หมวดหมู่"}}
ห้ามแต่งชื่อเกินความจำเป็น
"""

# ══════════════════════════════════════
# HELPERS
# ══════════════════════════════════════
def get_api_key():
    import os
    try: return str(st.secrets["GEMINI_API_KEY"]).strip()
    except: return os.environ.get("GEMINI_API_KEY", "").strip()

def gemini_classify(products, api_key):
    product_list = "\n".join(f"- {p}" for p in products)
    prompt = SYSTEM_PROMPT + "\n\nรายการสินค้า:\n" + product_list
    
    payload = json.dumps({
        "contents":[{"parts":[{"text":prompt}]}],
        "generationConfig":{"temperature":0.1, "maxOutputTokens":4000},
    }, ensure_ascii=False).encode("utf-8")
    
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={api_key}"
    req = urllib.request.Request(url, data=payload, headers={"Content-Type":"application/json; charset=utf-8"})
    
    with urllib.request.urlopen(req, timeout=60) as r:
        data = json.loads(r.read().decode("utf-8"))
    
    raw = re.sub(r"```json|```","",data["candidates"][0]["content"]["parts"][0]["text"]).strip()
    return json.loads(raw)

def load_excel(file):
    df = pd.read_excel(file, sheet_name="ใบเสร็จ")
    # สร้างคอลัมน์เก็บผลลัพธ์
    df["ชื่อสินค้า_ใหม่"] = df["ชื่อสินค้า"]
    df["ประเภทสินค้า"] = "Processed Food"
    return df

# ══════════════════════════════════════
# MAIN APP
# ══════════════════════════════════════
st.title("🛒 CJ Smart Scan Pro")
api_key = get_api_key()

uploaded = st.file_uploader("อัปโหลดไฟล์ Excel", type=["xlsx"])

if uploaded:
    if "df" not in st.session_state:
        st.session_state.df = load_excel(uploaded)
    
    df = st.session_state.df
    st.write(f"พบข้อมูล {len(df)} รายการ")
    
    if st.button("🤖 AI ทำความสะอาดและแยกหมวดหมู่ตาม PMA"):
        with st.spinner("กำลังวิเคราะห์..."):
            unique_prods = df["ชื่อสินค้า"].dropna().unique().tolist()
            # แบ่ง Batch ละ 50 รายการ
            for i in range(0, len(unique_prods), 50):
                batch = unique_prods[i:i+50]
                try:
                    ai_data = gemini_classify(batch, api_key)
                    for original, info in ai_data.items():
                        mask = df["ชื่อสินค้า"] == original
                        df.loc[mask, "ชื่อสินค้า_ใหม่"] = info["full_name"]
                        df.loc[mask, "ประเภทสินค้า"] = info["category"]
                except Exception as e:
                    st.error(f"Error: {e}")
                    break
            
            st.session_state.df = df
            st.rerun()

    if st.session_state.df is not None:
        st.dataframe(st.session_state.df)
        
        # ดาวน์โหลดไฟล์
        towrite = io.BytesIO()
        st.session_state.df.to_excel(towrite, index=False)
        towrite.seek(0)
        st.download_button("⬇️ ดาวน์โหลดไฟล์ที่ทำความสะอาดแล้ว", towrite, "CJ_Cleaned_Data.xlsx")
