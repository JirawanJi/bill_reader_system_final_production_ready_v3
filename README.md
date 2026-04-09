# Bill Reader System (Final PEA Template Version)

เวอร์ชันนี้ปรับ PEA exporter ให้ตรงกับเทมเพลต `FV60_08042026_015547.xlsx` แล้ว

## จุดเด่น
- Parser PEA ปรับตามบิลจริง
- Export PEA ออกคอลัมน์ตรงกับเทมเพลตจริง
- ใช้ไฟล์ `Cost_Center_2026_PROFIT_CENTER.xlsx` สำหรับ map Cost Center / Profit Center
- รองรับ PostgreSQL และ SQLite

## วิธีรัน
```bash
pip install -r requirements.txt
python app.py
```

## PEA Output
คอลัมน์หลักที่ตั้งค่าให้อัตโนมัติ:
- Company Code = 7590
- Vendor = 560014146
- Tax = X
- Tax Code = I3
- General Ledger = 41511101
- D.C = S
- Assignment = Vendor Invoice
- Payment Term = 0001
- Header Text = Provincial Electricity Authority
- Payment Method = D
- House Bank = CIT01
- Account ID = CI002

## ไฟล์ตัวอย่างสำหรับเทสต์
- PEA bill sample ใช้สาขา 83146
