import os
import re
from datetime import datetime

import pdfplumber
import fitz
import pytesseract
from PIL import Image


# =========================
# OCR + TEXT EXTRACTION
# =========================
def extract_text_from_pdf(filepath: str) -> str:
    texts = []

    try:
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                if page_text.strip():
                    texts.append(page_text)
    except Exception:
        pass

    try:
        doc = fitz.open(filepath)
        for page in doc:
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            try:
                text = pytesseract.image_to_string(
                    img,
                    lang="tha+eng",
                    config="--psm 6"
                )
            except Exception:
                text = pytesseract.image_to_string(img)

            if text.strip():
                texts.append(text)
    except Exception:
        pass

    return "\n".join(texts)


# =========================
# COMMON HELPERS
# =========================
def find_first(patterns, text, default=""):
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE | re.DOTALL)
        if match:
            return match.group(1).strip()
    return default


def normalize_amount(value: str) -> float:
    if not value:
        return 0.0
    value = value.replace(",", "").strip()
    value = re.sub(r"[^\d.]", "", value)
    try:
        return float(value)
    except Exception:
        return 0.0


def convert_thai_year(date_str: str) -> str:
    """
    แปลง date format ไทย -> สากล
      09/03/69   -> 09.03.2026
      09/03/2569 -> 09.03.2026
    """
    if not date_str:
        return ""

    m = re.match(r"^\s*(\d{1,2})/(\d{1,2})/(\d{2,4})\s*$", date_str)
    if not m:
        return ""

    d, mth, y = m.groups()
    d   = int(d)
    mth = int(mth)
    y   = int(y)

    if y < 100:
        y += 2500

    if y > 2400:
        y -= 543

    return f"{d:02d}.{mth:02d}.{y}"


def parse_ddmmyyyy(date_str: str):
    try:
        return datetime.strptime(date_str, "%d.%m.%Y")
    except Exception:
        return None


def normalize_store_id(store_id: str) -> str:
    if not store_id:
        return ""
    digits = re.sub(r"\D", "", str(store_id))
    if not digits:
        return ""
    if len(digits) == 4:
        return digits.zfill(5)
    if len(digits) == 5:
        return digits
    return digits


# =========================
# STORE ID
# =========================
def detect_store_id(text: str, filename: str = "") -> str:
    # 1) จากชื่อไฟล์ก่อน
    m = re.search(r"(?<!\d)(0\d{4}|8\d{4})(?!\d)", filename)
    if m:
        return normalize_store_id(m.group(1))

    # 2) จากท้ายบิล เช่น 01293 Feb'26 / 83226 Feb'26
    m = re.search(
        r"(?<!\d)(\d{4,5})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)'?\d{2}",
        text,
        re.IGNORECASE,
    )
    if m:
        return normalize_store_id(m.group(1))

    # 3) fallback ทั่วไป
    store_id = find_first(
        [
            r"(?<!\d)(0\d{4})(?!\d)",
            r"(?<!\d)(8\d{4})(?!\d)",
        ],
        text,
        "",
    )
    return normalize_store_id(store_id)


# =========================
# DATE EXTRACTION
# =========================
def extract_mea_meter_date(text: str, invoice: str = "") -> str:
    """ดึง Meter Reading Date สำหรับคอลัมน์ D / F"""
    date = find_first(
        [
            r"Meter\s*Reading\s*Date\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})",
            r"วันที่จดเลขอ่าน\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})",
            r"วันที่อ่านหน่วย\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})",
        ],
        text,
    )
    if date:
        return date

    if invoice:
        pattern = rf"{re.escape(invoice)}\s+([0-9]{{1,2}}/[0-9]{{1,2}}/[0-9]{{2,4}})"
        date = find_first([pattern], text)
        if date:
            return date

    date = find_first(
        [r"Invoice\s*No.*?Meter\s*Reading\s*Date.*?([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})"],
        text,
    )
    return date


def extract_mea_due_date(text: str) -> str:
    """
    ดึง Payment Due Date สำหรับคอลัมน์ AM

    โครงสร้าง OCR จริงจากบิล MEA (ทั้ง 2 ไฟล์ยืนยันแล้ว):
      บรรทัด N  : "...shiiutwytmwlusun 09/03/69"   <-- date อยู่ที่นี่
      บรรทัด N+1: "24515959697 27/02/69 ... Payment Due Date"  <-- keyword อยู่ที่นี่

    OCR ตัดบรรทัดให้ date อยู่ก่อน keyword 1 บรรทัด
    ดังนั้น logic ต้องมองย้อนกลับ (look-back) ไม่ใช่มองไปข้างหน้า

    กลยุทธ์ (เรียงตามความน่าเชื่อถือ):
    1) look-back  : หา line ที่มี keyword แล้วดึง date จากบรรทัดก่อนหน้า
    2) same-line  : กรณี OCR รวมบรรทัดเดียวกัน
    3) look-ahead : กรณี OCR ตัดแบบกลับกัน (สำรอง)
    4) inline Thai: ดึงจาก pattern "กำหนดชำระภายในวันที่ XX/XX/XX"
    """
    DATE_PATTERN = r"(\d{1,2}/\d{1,2}/\d{2,4})"

    KEYWORDS = [
        "payment due date",
        "กำหนดชำระภายในวันที่",
        "วันครบกำหนดชำระ",
        "ถึงวันที่รับชำระ",
    ]

    lines = text.splitlines()

    for i, line in enumerate(lines):
        line_lower = line.lower()

        if not any(kw in line_lower for kw in KEYWORDS):
            continue

        # กลยุทธ์ 1 (หลัก): date อยู่บรรทัดก่อนหน้า keyword
        for prev_line in reversed(lines[max(0, i - 3): i]):
            m = re.search(DATE_PATTERN, prev_line)
            if m:
                candidate = m.group(1)
                # กรองให้ได้เฉพาะ date ที่เป็น due date จริง
                # (ไม่ใช่ invoice/meter date ที่อาจปนใน header row เดียวกัน)
                # ใช้ตำแหน่ง: date ที่อยู่ท้ายบรรทัดมากกว่า ให้ความสำคัญกว่า
                dates_in_prev = re.findall(DATE_PATTERN, prev_line)
                if dates_in_prev:
                    # เอา date ที่อยู่ขวาสุดของบรรทัด (ท้ายบรรทัด = due date)
                    return dates_in_prev[-1]

        # กลยุทธ์ 2: date อยู่บรรทัดเดียวกับ keyword
        dates_same = re.findall(DATE_PATTERN, line)
        if dates_same:
            return dates_same[-1]

        # กลยุทธ์ 3: date อยู่บรรทัดถัดไป
        for next_line in lines[i + 1: i + 3]:
            m = re.search(DATE_PATTERN, next_line)
            if m:
                return m.group(1)

    # กลยุทธ์ 4: fallback pattern inline
    fallback = find_first(
        [
            r"กำหนดชำระภายในวันที่\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})",
            r"Payment\s*Due\s*Date\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})",
        ],
        text,
    )
    return fallback


# =========================
# AMOUNT EXTRACTION
# =========================
def extract_mea_amount(text: str) -> str:
    amount = find_first(
        [
            r"Amount\s*([\d,]+\.\d{2})",
            r"รวมเงินที่ต้องชำระทั้งสิ้น\s*([\d,]+\.\d{2})",
            r"รวมเงินค่าไฟฟ้าเดือนปัจจุบัน\s*([\d,]+\.\d{2})",
            r"รวมเงินทั้งสิ้น\s*[\(\w\s\)]*([\d,]+\.\d{2})",
            r"Payment\s*Due\s*Date.*?Amount\s*([\d,]+\.\d{2})",
        ],
        text,
        "",
    )
    if amount:
        return amount

    candidates = re.findall(r"\b\d{1,3}(?:,\d{3})*\.\d{2}\b", text)
    if candidates:
        try:
            return max(candidates, key=lambda x: float(x.replace(",", "")))
        except Exception:
            return candidates[-1]

    return "0.00"


# =========================
# TEXT / REF KEY
# =========================
def derive_month_year_from_invoice_date(invoice_date: str) -> str:
    dt = parse_ddmmyyyy(invoice_date)
    if not dt:
        return ""
    return dt.strftime("%b %Y")


def build_mea_text(store_id: str, invoice_date: str) -> str:
    month_year = derive_month_year_from_invoice_date(invoice_date)
    if store_id and month_year:
        return f"{store_id} - Electricity expense for {month_year}"
    if store_id:
        return f"{store_id} - Electricity expense"
    return "Electricity expense"


def build_mea_ref_key1(invoice_date: str) -> str:
    """
    D = 27.02.2026  ->  KB7189-01.26  (ใช้เดือนก่อนหน้า)
    """
    dt = parse_ddmmyyyy(invoice_date)
    if not dt:
        return "KB7189-01.26"

    month = dt.month - 1
    year  = dt.year
    if month == 0:
        month = 12
        year -= 1

    return f"KB7189-{month:02d}.{str(year)[-2:]}"


# =========================
# MAIN PARSER
# =========================
def parse_mea_pdf(filepath: str, mapping: dict, original_filename: str = "") -> dict:
    text = extract_text_from_pdf(filepath)

    store_id = detect_store_id(text, original_filename or os.path.basename(filepath))

    invoice = find_first(
        [
            r"\b(219\d{8})\b",
            r"\b(25\d{8,9})\b",
            r"\b(\d{10,11})\b",
        ],
        text,
    )

    amount_raw   = extract_mea_amount(text)
    meter_date   = extract_mea_meter_date(text, invoice)
    due_date_raw = extract_mea_due_date(text)   # <-- Payment Due Date จริง

    # validate due_date ก่อนใช้
    _due_converted = convert_thai_year(due_date_raw)
    if parse_ddmmyyyy(_due_converted):
        due_date = due_date_raw
    else:
        print(f"[WARN] due_date invalid or not found: {due_date_raw!r} -> blank")
        due_date = ""

    # meter_date fallback ถ้าไม่มีให้ใช้ due_date แทน
    if not meter_date:
        meter_date = due_date

    invoice_date  = convert_thai_year(meter_date)   # คอลัมน์ D
    posting_date  = invoice_date                     # คอลัมน์ F
    baseline_date = convert_thai_year(due_date)      # คอลัมน์ AM  ✅

    amount = normalize_amount(amount_raw)

    cost = (
        mapping.get(store_id, {})
        or mapping.get(store_id.lstrip("0"), {})
        or mapping.get(store_id.zfill(5), {})
    )

    text_desc = build_mea_text(store_id, invoice_date)
    ref_key1  = build_mea_ref_key1(invoice_date)

    print("====== MEA DEBUG ======")
    print(f"store_id      = {store_id}")
    print(f"invoice       = {invoice}")
    print(f"meter_date    = {meter_date}  (raw)")
    print(f"due_date_raw  = {due_date_raw}  (raw from Payment Due Date)")
    print(f"due_date      = {due_date}  (after validate)")
    print(f"invoice_date  = {invoice_date}  -> col D")
    print(f"posting_date  = {posting_date}  -> col F")
    print(f"baseline_date = {baseline_date}  -> col AM ✅")
    print(f"amount_raw    = {amount_raw}")
    print(f"amount        = {amount}")
    print(f"cost          = {cost}")
    print(f"text_desc     = {text_desc}")
    print(f"ref_key1      = {ref_key1}")
    print("=======================")

    return {
        "company_code":   "7590",
        "document_type":  "KR",
        "vendor":         "560014090",

        "invoice_date":   invoice_date,    # D
        "posting_date":   posting_date,    # F
        "baseline_date":  baseline_date,   # AM ✅

        "reference":      invoice,
        "currency":       "THB",
        "amount":         amount,

        "tax_code":       "I3",
        "business_place": "7590",
        "section":        "",

        "text":           text_desc,
        "assignment":     "Vendor Invoice",

        "cost_center":    cost.get("cost_center", ""),
        "profit_center":  cost.get("profit_center", ""),

        "payment_term":   "0001",
        "header_text":    "Metropolitan Electricity Authority",

        "ref_key1":       ref_key1,
        "ref_key2":       "",
        "ref_key3":       "",

        "payment_method": "K",
        "partner_bank":   "",
        "house_bank":     "KBK04",
        "account_id":     "KB189",

        "store_id":       store_id,
        "filename":       original_filename or os.path.basename(filepath),
    }