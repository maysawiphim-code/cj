# -*- coding: utf-8 -*-
"""
CJ Smart Scan — Python Edition
================================
วิเคราะห์ยอดลูกค้าสะสมรายสาขา & จำแนกสินค้าตาม PMA 5 กลุ่ม

Logic ตรงกับเวอร์ชันเว็บ (lib/cjConstants.js + lib/cjClassifier.js)

ติดตั้ง:
    pip install openpyxl google-generativeai

การใช้งาน:
    from cj_smartscan import classifyRuleOnly, classifyWithAI, parseExcelData, \
        getCustomerAccumulate, getBranchCustomerTotal, makeSummary
"""

import re
import json
from collections import defaultdict
from datetime import datetime

# ═══════════════════════════════════════════════════
# 1. CONSTANTS — PMA 5 กลุ่ม + คีย์เวิร์ด
# ═══════════════════════════════════════════════════

CATEGORIES = {
    "Fresh Food": {
        "icon": "🥩", "color": "#E53935", "hex": "E53935",
        "sub": [
            "CHILLED BREAD", "APPETIZER", "FRUIT", "COUNTER DRINK", "FOOD PLACE",
            "SANDWICH", "BURGER", "สลัด", "MEAL BOX", "FROZEN", "SAUSAGE", "GRILLED",
            "WARMED", "RETORT", "PASTEURIZED", "PACKAGE BAKERY", "SANDWICH BREAD",
            "VARIETY BAKERY", "CPG SYNERGY", "READY TO COOK", "RTC"
        ],
    },
    "Non Food": {
        "icon": "🏠", "color": "#1565C0", "hex": "1565C0",
        "sub": [
            "CIGARETTE", "BOOKS", "ENTERTAINMENT", "IT Device", "PERSONAL CARE",
            "HOUSEWARE", "STATIONERY", "SANITARY", "HOUSEHOLD", "ELECTRONIC",
            "IT APPLIANCE", "HERBAL", "เทศกาล", "MAGAZINE", "NEWSPAPER"
        ],
    },
    "Packaged Beverage": {
        "icon": "🥤", "color": "#00897B", "hex": "00897B",
        "sub": [
            "BEER", "ALCOHOL", "LIQUOR", "UHT MILK", "NON-CARBONATED", "CARBONATED",
            "CSD", "ICE CREAM", "NOVELTIES", "ICE", "ENERGY", "SPORT DRINK"
        ],
    },
    "Processed Food": {
        "icon": "🍜", "color": "#F57C00", "hex": "F57C00",
        "sub": [
            "COOKING", "CANNED", "INSTANT FOODS", "PACKAGED FOODS", "CONFECTIONERY",
            "SNACKS", "THAI SNACK", "DRY FRUIT"
        ],
    },
    "Special Business": {
        "icon": "⭐", "color": "#7B1FA2", "hex": "7B1FA2",
        "sub": [
            "HEALTH CARE", "MEDICINE", "SYNERGY PROJECT", "7 Service", "WELLNESS",
            "DRUG", "FRESH BAKERY", "COMMISSION", "BELLINEE", "KUDSAN", "VEGETABLE",
            "BEVERAGE", "CATALOG", "SUPPLY", "HOT SERVED", "SOCIAL WELFARE",
            "HOME & LIVING", "FASHION"
        ],
    },
}

RULE_KW = {
    "Fresh Food": [
        "bao", "bac", "แซนด์วิช", "ไส้กรอก", "ขนมปัง", "เบเกอรี", "sandwich",
        "sausage", "ข้าวกล่อง", "ไข่", "salad", "สลัด", "นม pasteur",
        "โยเกิร์ต", "ไก่", "หมู", "เนื้อ", "ทูน่า", "ปลา", "อาหาร", "meal",
    ],
    "Non Food": [
        "ยาสีฟัน", "แชมพู", "สบู่", "ครีม", "ผ้าอนามัย", "ทิชชู", "ถุงขยะ",
        "ผงซักฟอก", "น้ำยา", "บุหรี่", "cigarette", "ไฟแช็ก", "ถ่าน",
        "แปรง", "โฟม", "โลชั่น", "เทปลบ", "ปากกา", "สมุด", "personal care",
        "household", "sanitary", "houseware", "stationery", "electronic",
    ],
    "Packaged Beverage": [
        "น้ำดื่ม", "น้ำแร่", "โค้ก", "เป๊ปซี่", "สไปรท์", "คาราบาว",
        "เรดบูล", "ไมโล", "นมuht", "uht", "เบียร์", "beer", "วิสกี้",
        "สุรา", "ไอศกรีม", "ice cream", "น้ำผลไม้", "ชาเขียว", "โอเลี้ยง",
        "energy drink", "soft drink", "mineral", "โซดา", "csd",
    ],
    "Processed Food": [
        "มาม่า", "บะหมี่", "instant", "ขนมกรุบ", "มันฝรั่ง", "ลูกอม",
        "ช็อกโกแลต", "ถั่ว", "เมล็ด", "snack", "confectionery",
        "canned", "กระป๋อง", "ข้าวสาร", "เส้น", "pasta",
    ],
    "Special Business": [
        "เติมเงิน", "จ่ายบิล", "ส่วนลด", "discount", "บริการ",
        "ยา", "vitamin", "วิตามิน", "health", "wellness", "อาหารเสริม",
        "bellinee", "kudsan", "social welfare", "7-service",
    ],
}

VALID_CATS = set(CATEGORIES.keys())

RANK_BG = [
    "#FFD700", "#C0C0C0", "#CD7F32", "#FF6B6B", "#FF9F43",
    "#48DBFB", "#1DD1A1", "#A29BFE", "#FD79A8", "#636E72",
]
RANK_MEDAL = ["🥇", "🥈", "🥉", "4", "5", "6", "7", "8", "9", "10"]

SYSTEM_PROMPT = """คุณคือระบบจำแนกประเภทสินค้าของ CJ Express (CJ More)
จงวิเคราะห์จากชื่อสินค้า และกำหนดให้อยู่ใน 1 ใน 5 กลุ่ม PMA ต่อไปนี้เท่านั้น:

1. Fresh Food
2. Non Food
3. Packaged Beverage
4. Processed Food
5. Special Business

กฎสำคัญ:
- ตอบเป็น JSON object เพียงอย่างเดียว ไม่มีคำอธิบายเพิ่ม ไม่มี backticks
- คีย์ = ชื่อสินค้า (ตามที่ส่งมาเป๊ะ), ค่า = 1 ใน 5 กลุ่มด้านบน
- หากไม่แน่ใจ ให้เลือก "Processed Food"
- ห้ามสร้างรายการใหม่ที่ไม่ได้ส่งมา

ตัวอย่าง:
{ "แซนด์วิชทูน่า": "Fresh Food", "ยาสีฟันคอลเกต": "Non Food", "โค้ก 1.25L": "Packaged Beverage", "มันฝรั่งทอดรสบาร์บิคิว": "Processed Food", "บริการเติมเงินมือถือ": "Special Business" }"""


# ═══════════════════════════════════════════════════
# 2. CLASSIFICATION ENGINE
# ═══════════════════════════════════════════════════

def rule_classify(name):
    """จำแนกด้วย keyword matching → คืนชื่อกลุ่ม หรือ "" ถ้าไม่ตรง"""
    n = str(name).lower()
    for cat, kws in RULE_KW.items():
        if any(k.lower() in n for k in kws):
            return cat
    return ""


def classify_rule_only(products):
    """จำแนกเฉพาะ rule-based; รายการที่ไม่ตรง → 'Processed Food'"""
    m = {}
    for p in products:
        cat = rule_classify(p)
        m[p] = cat or "Processed Food"
    return m


# ═══════════════════════════════════════════════════
# 3. SAFE JSON PARSER + NORMALIZE
# ═══════════════════════════════════════════════════

def parse_json_safe(text):
    """แปลง string ที่อาจมี code block / trailing comma → dict"""
    cleaned = text.strip()
    # ลบ code block markers
    cleaned = re.sub(r'^```(?:json)?\s*', '', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'\s*```$', '', cleaned, flags=re.IGNORECASE).strip()
    # ดึง JSON object แรก
    m = re.search(r'\{[\s\S]*\}', cleaned)
    if m:
        cleaned = m.group(0)
    # แก้ trailing comma
    cleaned = re.sub(r',(\s*[}\]])', r'\1', cleaned)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        raise ValueError(f"ไม่สามารถแปลงผลลัพธ์เป็น JSON ได้: {text[:400]}")


def normalize_mapping(result):
    """บังคับให้ทุกค่าเป็น 1 ใน 5 กลุ่มที่ถูกต้อง"""
    out = {}
    for k, v in result.items():
        cat = str(v).strip()
        out[str(k)] = cat if cat in VALID_CATS else "Processed Food"
    return out


# ═══════════════════════════════════════════════════
# 4. AI CLASSIFY (Gemini)
# ═══════════════════════════════════════════════════

def gemini_classify(products, api_key=None):
    """
    เรียก Gemini API เพื่อจำแนกสินค้าเป็น batch
    คืน dict {ชื่อสินค้า: กลุ่ม}
    """
    try:
        import google.generativeai as genai
    except ImportError:
        raise ImportError("ติดตั้ง: pip install google-generativeai")

    import os
    key = api_key or os.environ.get("GEMINI_API_KEY")
    if not key:
        raise ValueError("ตั้ง GEMINI_API_KEY หรือส่ง api_key")

    genai.configure(api_key=key)
    model = genai.GenerativeModel(
        "gemini-1.5-flash",
        system_instruction=SYSTEM_PROMPT,
        generation_config={"response_mime_type": "application/json"},
    )

    product_list = "\n".join(f"- {p}" for p in products)
    prompt = SYSTEM_PROMPT + "\n\nจำแนกสินค้าเหล่านี้:\n" + product_list

    response = model.generate_content(prompt)
    raw = response.text
    raw_obj = parse_json_safe(raw) if isinstance(raw, str) else raw
    result = normalize_mapping(raw_obj)

    # เติมรายการที่หายไปด้วย fallback
    for p in products:
        if p not in result:
            result[p] = "Processed Food"
    return result


def classify_with_ai(products, current_map=None, on_progress=None, api_key=None,
                     batch_size=30):
    """
    จำแนกด้วย AI สำหรับรายการที่ rule-based ไม่ตรับ
    รวมผลเข้ากับ current_map (ไม่เขียนทับค่าที่ดีอยู่แล้ว)
    """
    if current_map is None:
        current_map = {}
    result = dict(current_map)

    need_ai = [p for p in products if not rule_classify(p)]
    if not need_ai:
        return result

    batches = [need_ai[i:i + batch_size] for i in range(0, len(need_ai), batch_size)]

    for i, batch in enumerate(batches):
        if on_progress:
            on_progress(min((i + 1) * batch_size, len(need_ai)), len(need_ai),
                        i + 1, len(batches))
        try:
            ai_map = gemini_classify(batch, api_key=api_key)
            for k, v in ai_map.items():
                if not result.get(k) or result[k] == "" or result[k] == "Processed Food":
                    result[k] = v
        except Exception as e:
            err = str(e).lower()
            if "429" in err or "quota" in err:
                raise RuntimeError("rate_limit")
            print(f"[WARN] Batch {i + 1} error: {e}")

    return result


# ═══════════════════════════════════════════════════
# 5. RECEIPT & EXCEL PARSING
# ═══════════════════════════════════════════════════

def parse_receipt_no(receipt_no):
    """แยกเลขเครื่อง + เลขทรานแซกชันจาก 'N0001-1234'"""
    m = re.match(r"(N\d+)-(\d+)", str(receipt_no))
    if m:
        return {"machine": m.group(1), "trans_no": int(m.group(2))}
    return {"machine": None, "trans_no": None}


def parse_excel_data(rows):
    """
    แปลง list[dict] จาก Excel ให้มีฟิลด์ _machine, _trans_no, _parsed_date, _month
    """
    out = []
    for row in rows:
        receipt = parse_receipt_no(row.get("เลขที่ใบเสร็จ", ""))
        date_str = row.get("วันที่", "")
        parsed_date = None
        month_label = None

        if date_str:
            parts = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", str(date_str))
            if parts:
                d, m, y = int(parts.group(1)), int(parts.group(2)), int(parts.group(3))
                parsed_date = datetime(y, m, d)
                thai_year = y + 543
                month_label = f"{parsed_date.strftime('%b')} {thai_year}"

        out.append({
            **row,
            "_machine": receipt["machine"],
            "_trans_no": receipt["trans_no"],
            "_parsed_date": parsed_date,
            "_month": month_label,
        })
    return out


# ═══════════════════════════════════════════════════
# 6. AGGREGATION & REPORTING
# ═══════════════════════════════════════════════════

def get_customer_accumulate(rows):
    """ยอดลูกค้าสะสม = max trans_no ของแต่ละเครื่องต่อสาขา"""
    machine_max = {}
    for row in rows:
        if not row.get("_machine") or row.get("_trans_no") is None:
            continue
        branch = row.get("รหัสสาขา")
        if not branch:
            continue
        key = f"{branch}__{row['_machine']}"
        machine_max[key] = max(machine_max.get(key, 0), row["_trans_no"])

    return [
        {"branch": k.split("__")[0], "machine": k.split("__")[1], "customer_count": v}
        for k, v in machine_max.items()
    ]


def get_branch_customer_total(rows):
    """รวมยอดลูกค้าสะสมตามสาขา → เรียงจากมากไปน้อย"""
    mc = get_customer_accumulate(rows)
    totals = defaultdict(int)
    for item in mc:
        totals[item["branch"]] += item["customer_count"]
    return sorted(
        [{"branch": b, "total": t} for b, t in totals.items()],
        key=lambda x: x["total"], reverse=True
    )


def make_summary(rows, cat_map):
    """สรุปจำนวน + ยอดรวมตามกลุ่ม PMA"""
    summary = defaultdict(lambda: {"count": 0, "total": 0.0})
    for row in rows:
        name = row.get("ชื่อสินค้า")
        if not name:
            continue
        cat = cat_map.get(name, "Processed Food")
        summary[cat]["count"] += 1
        summary[cat]["total"] += float(row.get("ยอดรวมสินค้า") or 0)

    return sorted(
        [{"category": c, "count": d["count"], "total": d["total"]}
         for c, d in summary.items()],
        key=lambda x: x["count"], reverse=True
    )


# ═══════════════════════════════════════════════════
# 7. EXCEL READER (openpyxl)
# ═══════════════════════════════════════════════════

def read_excel_rows(file_path):
    """อ่านไฟล์ .xlsx → list[dict] (ใช้ header row เป็น key)"""
    try:
        from openpyxl import load_workbook
    except ImportError:
        raise ImportError("ติดตั้ง: pip install openpyxl")

    wb = load_workbook(file_path, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    headers = [str(h).strip() if h else f"col_{i}" for i, h in enumerate(next(rows_iter))]

    out = []
    for row in rows_iter:
        if all(v is None for v in row):
            continue
        out.append({headers[i]: row[i] for i in range(len(headers)) if i < len(row)})
    return out


# ═══════════════════════════════════════════════════
# 8. CLI ENTRY POINT
# ═══════════════════════════════════════════════════

def main():
    import argparse
    parser = argparse.ArgumentParser(description="CJ Smart Scan — วิเคราะห์ยอดลูกค้า + PMA 5 กลุ่ม")
    parser.add_argument("file", help="ไฟล์ .xlsx")
    parser.add_argument("--ai", action="store_true", help="เปิดใช้ AI classify")
    parser.add_argument("--output", "-o", default=None, help="บันทึกผลลัพธ์เป็น JSON")
    args = parser.parse_args()

    print("📂 อ่านไฟล์:", args.file)
    raw_rows = read_excel_rows(args.file)
    rows = parse_excel_data(raw_rows)
    print(f"✅ พบ {len(rows):,} รายการ จาก {len(set(r.get('รหัสสาขา','') for r in rows))} สาขา")

    products = list({r["ชื่อสินค้า"] for r in rows if r.get("ชื่อสินค้า")})
    print(f"🏷️ สินค้าไม่ซ้ำ: {len(products):,}")

    print("\n⚡ จำแนกด้วย Rule-based...")
    cat_map = classify_rule_only(products)
    need_ai = [p for p in products if not rule_classify(p)]
    print(f"   ตรง rule: {len(products) - len(need_ai):,} | ต้องใช้ AI: {len(need_ai):,}")

    if args.ai and need_ai:
        print("\n🤖 จำแนกเพิ่มด้วย AI...")
        cat_map = classify_with_ai(products, cat_map, on_progress=lambda d, t, b, n: print(f"   🤖 {d}/{t} (batch {b}/{n})"))
        print("✅ AI จำแนกเสร็จ!")

    print("\n" + "=" * 50)
    print("📊 สรุปตาม PMA 5 กลุ่ม")
    print("=" * 50)
    summary = make_summary(rows, cat_map)
    for s in summary:
        print(f"  {s['category']:20s} | จำนวน: {s['count']:>6,} | ฿{s['total']:>12,.2f}")

    print("\n" + "=" * 50)
    print("👥 ยอดลูกค้าสะสมรายสาขา (Top 10)")
    print("=" * 50)
    branch_totals = get_branch_customer_total(rows)
    for i, b in enumerate(branch_totals[:10]):
        medal = RANK_MEDAL[i] if i < len(RANK_MEDAL) else str(i + 1)
        print(f"  {medal} {b['branch']:>8s}  →  {b['total']:>8,}")

    total_customers = sum(b["total"] for b in branch_totals)
    print(f"\n  📈 ยอดลูกค้าสะสมรวมทุกสาขา: {total_customers:,}")

    machine_data = get_customer_accumulate(rows)
    print("\n" + "=" * 50)
    print("🖨️ ยอดลูกค้าสะสมแยกตามเครื่อง (Top 5 สาขา)")
    print("=" * 50)
    for b in branch_totals[:5]:
        machines = [m for m in machine_data if m["branch"] == b["branch"]]
        machines.sort(key=lambda x: x["customer_count"], reverse=True)
        print(f"\n  🏪 สาขา {b['branch']} (รวม {b['total']:,})")
        for m in machines[:5]:
            print(f"     {m['machine']:>8s}  →  {m['customer_count']:>6,}")

    if args.output:
        result = {
            "summary": summary,
            "branch_totals": branch_totals,
            "machine_data": machine_data,
            "cat_map": cat_map,
        }
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print(f"\n💾 บันทึกผลลัพธ์ที่: {args.output}")


if __name__ == "__main__":
    main()
