import os
import re
from datetime import datetime, date, timedelta

import pdfplumber
import fitz
import pytesseract
from PIL import Image


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
            page_text = page.get_text("text") or ""
            if page_text.strip():
                texts.append(page_text)

            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            try:
                ocr_text = pytesseract.image_to_string(
                    img, lang="tha+eng", config="--psm 6"
                )
            except Exception:
                ocr_text = pytesseract.image_to_string(img)
            if ocr_text.strip():
                texts.append(ocr_text)
    except Exception:
        pass
    return "\n".join(texts)


def find_first(patterns, text, default=""):
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE | re.DOTALL)
        if match:
            return match.group(1).strip()
    return default


def normalize_amount(value: str) -> float:
    if not value:
        return 0.0
    cleaned = normalize_ocr_digits(value).replace(",", "").replace(" ", "")
    cleaned = re.sub(r"[^\d.]", "", cleaned)
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def normalize_ocr_digits(text: str) -> str:
    if not text:
        return ""
    # Thai digits + common OCR confusions
    table = str.maketrans({
        "๐": "0", "๑": "1", "๒": "2", "๓": "3", "๔": "4",
        "๕": "5", "๖": "6", "๗": "7", "๘": "8", "๙": "9",
        "O": "0", "o": "0", "D": "0", "Q": "0",
        "I": "1", "l": "1", "|": "1",
        "S": "5", "s": "5",
        "B": "8",
    })
    return text.translate(table)


def convert_thai_year(date_str: str) -> str:
    if not date_str:
        return ""

    m = re.match(r"^\s*(\d{1,2})/(\d{1,2})/(\d{2,4})\s*$", date_str)
    if not m:
        return ""

    d, mth, y = m.groups()
    d = int(d)
    mth = int(mth)
    y = int(y)

    if y < 100:
        y = 2500 + y

    if y > 2400:
        y -= 543

    return f"{d:02d}.{mth:02d}.{y}"


def convert_thai_text_date(date_str: str) -> str:
    if not date_str:
        return ""

    thai_months = {
        "มกราคม": 1,
        "กุมภาพันธ์": 2,
        "มีนาคม": 3,
        "เมษายน": 4,
        "พฤษภาคม": 5,
        "มิถุนายน": 6,
        "กรกฎาคม": 7,
        "สิงหาคม": 8,
        "กันยายน": 9,
        "ตุลาคม": 10,
        "พฤศจิกายน": 11,
        "ธันวาคม": 12,
    }

    m = re.match(r"^\s*(\d{1,2})\s*([ก-๙]+)\s*(\d{4})\s*$", date_str)
    if not m:
        return ""

    d = int(m.group(1))
    thai_month = m.group(2).strip()
    y = int(m.group(3))

    month = thai_months.get(thai_month)
    if not month:
        return ""

    if y > 2400:
        y -= 543

    return f"{d:02d}.{month:02d}.{y}"


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


def detect_store_id(text: str, filename: str = "") -> str:
    # 1) จากชื่อไฟล์ก่อน เช่น 01266 - ... / 83224 - ...
    m = re.search(r"(?<!\d)(0\d{4}|8\d{4})(?!\d)", filename)
    if m:
        return normalize_store_id(m.group(1))

    # 2) จากท้ายบิล เช่น 1266 Mar'26 / 83224 Mar'26 / 83226 Feb'26
    m = re.search(
        r"(?<!\d)(\d{4,5})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)'?\d{2}",
        text,
        re.IGNORECASE,
    )
    if m:
        return normalize_store_id(m.group(1))

    # 3) explicit field
    store_id = find_first(
        [
            r"Store\s*ID\s*[:\-]?\s*(0\d{4}|8\d{4}|\d{4,5})",
            r"Branch\s*[:\-]?\s*(0\d{4}|8\d{4}|\d{4,5})",
            r"รหัสสาขา\s*[:\-]?\s*(0\d{4}|8\d{4}|\d{4,5})",
        ],
        text,
        default="",
    )
    if store_id:
        return normalize_store_id(store_id)

    # 4) fallback เลข 5 หลักในเอกสาร
    store_id = find_first(
        [
            r"(?<!\d)(0\d{4})(?!\d)",
            r"(?<!\d)(8\d{4})(?!\d)",
        ],
        text,
        default="",
    )

    return normalize_store_id(store_id)


def derive_bill_month_year_from_text(text: str, fallback_date: str = "") -> str:
    m = re.search(
        r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[\'\s\-]?(\d{2})\b",
        text,
        re.IGNORECASE,
    )
    if m:
        mon = m.group(1).title()
        yy = int(m.group(2))
        return f"{mon} {2000 + yy}"

    dt = parse_ddmmyyyy(fallback_date)
    if dt:
        return dt.strftime("%b %Y")

    return ""


def extract_pea_invoice(text: str) -> str:
    invoice = find_first(
        [
            r"Invoice\s*no\.?\s*([0-9]+)",
            r"เลขที่ใบแจ้งค่าไฟฟ้า\s*([0-9]+)",
            r"\b(869\d{7})\b",
            r"\b(8788\d{6})\b",
            r"\b(\d{10,12})\b",
        ],
        text,
        default="",
    )
    invoice = normalize_ocr_digits(invoice)
    invoice = re.sub(r"\D", "", invoice)
    return invoice


def extract_pea_amount(text: str) -> str:
    normalized_text = normalize_ocr_digits(text)
    amount_pattern = r"(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2}"
    lines = [ln.strip() for ln in normalized_text.splitlines() if ln.strip()]

    def parse_amount_token(token: str) -> float:
        token = normalize_ocr_digits(token)
        token = re.sub(r"[^\d.,]", "", token)
        if not token:
            return 0.0
        # Fix malformed OCR token like ".99.013.05" -> "99013.05"
        if token.count(".") >= 2:
            last_dot = token.rfind(".")
            int_part = re.sub(r"\D", "", token[:last_dot])
            dec_part = re.sub(r"\D", "", token[last_dot + 1:])
            if int_part and len(dec_part) >= 2:
                token = f"{int_part}.{dec_part[:2]}"
        else:
            token = token.replace(",", "")
        return normalize_amount(token)

    def line_amounts(line: str):
        amounts = [parse_amount_token(x) for x in re.findall(amount_pattern, line)]
        # include malformed dotted values that standard pattern splits incorrectly
        malformed_tokens = re.findall(r"[.,]?\d[\d.,]*\.\d{2}", line)
        amounts.extend(parse_amount_token(x) for x in malformed_tokens)
        return [x for x in amounts if x > 0]

    current_month = []
    total_lines = []

    for line in lines:
        line_l = line.lower()
        amounts = line_amounts(line)
        if not amounts:
            continue

        # Most reliable: current-month total line (allow OCR distortion)
        if ("รวมเงิน" in line and ("เดือนปัจจุบ" in line or "เดือนบัจจุบ" in line)):
            current_month.extend(amounts)
            continue

        # Other total-like lines
        if (
            "grand total" in line_l
            or "รวมเงินทั้งสิ้น" in line
            or "รวมเงินค่าไฟฟ้า" in line
            or "sub total" in line_l
            or "total" in line_l
        ):
            total_lines.extend(amounts)

    if current_month:
        # Prefer realistic cap (OCR can prepend one extra digit)
        realistic = [x for x in current_month if x <= 300000]
        return f"{(max(realistic) if realistic else max(current_month)):,.2f}"

    if total_lines:
        realistic = [x for x in total_lines if x <= 300000]
        pick = max(realistic) if realistic else max(total_lines)
        # If still suspiciously tiny, try global max as fallback
        if pick < 500:
            all_amounts = [normalize_amount(x) for x in re.findall(amount_pattern, normalized_text) if normalize_amount(x) > 0]
            if all_amounts:
                realistic_all = [x for x in all_amounts if x <= 300000]
                pick = max(realistic_all) if realistic_all else max(all_amounts)
        return f"{pick:,.2f}"

    # 3) global monetary fallback
    all_amounts = [normalize_amount(x) for x in re.findall(amount_pattern, normalized_text) if normalize_amount(x) > 0]
    if all_amounts:
        realistic = [x for x in all_amounts if x <= 300000]
        return f"{(max(realistic) if realistic else max(all_amounts)):,.2f}"

    return "0.00"


def extract_pea_meter_reading_date(text: str, invoice: str = "") -> str:
    normalized_text = normalize_ocr_digits(text)
    # 1) แบบมี keyword
    date_raw = find_first(
        [
            r"Meter\s*Reading\s*Date\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})",
            r"วันที่อ่านหน่วย\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})",
        ],
        normalized_text,
        default="",
    )
    if date_raw:
        return date_raw

    # 2) จากตารางจริง เช่น ... 3224 28/03/2569 03/2569 ...
    date_raw = find_first(
        [
            r"\b\d{4}\s+([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})\s+\d{2}/\d{4}\b",
            r"\b\d{4}\s+([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})\s+\d{2}/\d{2}\b",
        ],
        normalized_text,
        default="",
    )
    if date_raw:
        return date_raw

    # 3) fallback จากบรรทัดที่มี invoice
    if invoice:
        date_raw = find_first(
            [
                rf"{re.escape(invoice)}.*?([0-9]{{1,2}}/[0-9]{{1,2}}/[0-9]{{2,4}})"
            ],
            normalized_text,
            default="",
        )
        if date_raw:
            return date_raw

    return ""


def extract_pea_document_date(text: str) -> str:
    normalized_text = normalize_ocr_digits(text)
    return find_first(
        [
            r"Document\s*Date\s*:\s*([0-9\/]+)",
            r"Date\s*:\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})",
        ],
        normalized_text,
        default="",
    )


def extract_pea_due_date(text: str) -> str:
    normalized_text = normalize_ocr_digits(text)
    return find_first(
        [
            r"Due\s*Date\s*([0-9]{1,2}\s+[^\s]+\s+[0-9]{4})",
            r"วันที่ครบกำหนดค่าไฟฟ้าเดือนปัจจุบัน.*?([0-9]{1,2}\s+[^\s]+\s+[0-9]{4})",
            r"วันที่ครบกําหนดค่าไฟฟ้าเดือนปัจจุบัน.*?([0-9]{1,2}\s*[^\s]+\s*[0-9]{4})",
        ],
        normalized_text,
        default="",
    )


def fourth_thursday_of_month(dt: date) -> date:
    first_day = dt.replace(day=1)
    days_until_thursday = (3 - first_day.weekday()) % 7
    first_thursday = first_day + timedelta(days=days_until_thursday)
    fourth_thursday = first_thursday + timedelta(weeks=3)
    return fourth_thursday


def build_pea_baseline_date(store_id: str, due_date_text: str, invoice_date: str) -> str:
    # ใช้ due_date เหมือนกันทุก store (รวมถึง 01266)
    return convert_thai_text_date(due_date_text)


def build_pea_ref_key1_from_baseline(baseline_date: str) -> str:
    dt = parse_ddmmyyyy(baseline_date)
    if not dt:
        return ""
    return f"ll-{dt.strftime('%d.%m.%y')}"


def parse_pea_pdf(filepath: str, mapping: dict, original_filename: str = "") -> dict:
    text = extract_text_from_pdf(filepath)
    filename_for_detect = original_filename or os.path.basename(filepath)

    store_id = detect_store_id(text, filename_for_detect)
    invoice = extract_pea_invoice(text)
    amount_raw = extract_pea_amount(text)

    meter_reading_date_raw = extract_pea_meter_reading_date(text, invoice)
    document_date_raw = extract_pea_document_date(text)
    due_date_thai_raw = extract_pea_due_date(text)

    invoice_date = convert_thai_year(meter_reading_date_raw)
    if not invoice_date:
        invoice_date = convert_thai_year(document_date_raw)

    posting_date = invoice_date
    baseline_date = build_pea_baseline_date(store_id, due_date_thai_raw, invoice_date)

    amount = normalize_amount(amount_raw)
    bill_month_year = derive_bill_month_year_from_text(text, invoice_date)

    cost = (
        mapping.get(store_id, {})
        or mapping.get(store_id.lstrip("0"), {})
        or mapping.get(store_id.zfill(5), {})
    )

    text_line = (
        f"{store_id} - Electricity expense for {bill_month_year}"
        if store_id
        else f"Electricity expense for {bill_month_year}"
    )
    ref_key1 = build_pea_ref_key1_from_baseline(baseline_date)

    print("====== PEA DEBUG ======")
    print("store_id =", store_id)
    print("invoice =", invoice)
    print("meter_reading_date_raw =", meter_reading_date_raw)
    print("document_date_raw =", document_date_raw)
    print("invoice_date =", invoice_date)
    print("posting_date =", posting_date)
    print("baseline_date =", baseline_date)
    print("cost =", cost)
    print("text_line =", text_line)
    print("=======================")

    return {
        "company_code": "7590",
        "document_type": "KR",
        "vendor": "560014146",

        "invoice_date": invoice_date,
        "posting_date": posting_date,
        "baseline_date": baseline_date,

        "reference": invoice,
        "currency": "THB",
        "amount": amount,

        "tax_code": "I3",
        "business_place": "7590",
        "section": "",

        "text": text_line,
        "assignment": "",

        "cost_center": cost.get("cost_center", ""),
        "profit_center": cost.get("profit_center", ""),

        "payment_term": "0001",
        "header_text": "Provincial Electricity Authority",
        "ref_key1": ref_key1,

        "payment_method": "D",
        "house_bank": "CIT01",
        "account_id": "CI002",

        "store_id": store_id,
        "filename": original_filename or os.path.basename(filepath),
        "raw_text_preview": text[:2000],
    }
