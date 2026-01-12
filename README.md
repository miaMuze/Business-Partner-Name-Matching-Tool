# โปรแกรมเช็คและจับคู่ชื่อ Business Partner (BP Check)

โปรแกรมนี้ใช้สำหรับจับคู่ชื่อลูกค้าระหว่างระบบ Oracle และ SAP โดยใช้เทคนิค Fuzzy Matching เพื่อหาชื่อที่ใกล้เคียงกัน แม้จะมีการสะกดที่แตกต่างกัน

## คุณสมบัติ

- จับคู่ชื่อจากข้อมูล Oracle กับ SAP
- แสดงผลลัพธ์ 5 อันดับที่ใกล้เคียงที่สุด
- มีคะแนนความเหมือน (Score) สำหรับแต่ละผลลัพธ์
- ทำความสะอาดชื่อก่อนเปรียบเทียบ (ลบคำนำหน้า, บริษัท, จำกัด, ฯลฯ)
- แสดงความคืบหน้าขณะทำงาน

## ข้อกำหนดระบบ

- Python 3.7 หรือสูงกว่า
- ไฟล์ Excel ข้อมูลต้นทาง (`data_migration.xlsx`)

## การติดตั้ง

### 1. ติดตั้ง Python

ตรวจสอบว่าคุณมี Python ติดตั้งอยู่แล้วหรือไม่:

```bash
python3 --version
```

หากยังไม่มี ดาวน์โหลดและติดตั้งจาก [python.org](https://www.python.org/)

### 2. สร้าง Virtual Environment

```bash
python3 -m venv .venv
```

### 3. เปิดใช้งาน Virtual Environment

**บน macOS/Linux:**
```bash
source .venv/bin/activate
```

**บน Windows:**
```bash
.venv\Scripts\activate
```

### 4. ติดตั้ง Dependencies

```bash
pip install pandas thefuzz openpyxl python-Levenshtein
```

## การเตรียมข้อมูล

สร้างไฟล์ Excel ชื่อ `data_migration.xlsx` ในโฟลเดอร์เดียวกับโปรแกรม โดยมี 2 Sheets:

### Sheet 1: Oracle
มีคอลัมน์ดังนี้:
- `ID` - รหัสลูกค้าใน Oracle
- `Name1` - ชื่อส่วนที่ 1
- `Name2` - ชื่อส่วนที่ 2 (ถ้ามี)

### Sheet 2: SAP
มีคอลัมน์ดังนี้:
- `BP_Number` - รหัส Business Partner ใน SAP
- `Name1` - ชื่อส่วนที่ 1
- `Name2` - ชื่อส่วนที่ 2 (ถ้ามี)

## การรันโปรแกรม

1. เปิด Terminal/Command Prompt
2. ไปที่โฟลเดอร์โปรแกรม:
   ```bash
   cd /path/to/BP_CHECK
   ```

3. เปิดใช้งาน Virtual Environment (ถ้ายังไม่ได้เปิด):
   ```bash
   source .venv/bin/activate
   ```

4. รันโปรแกรม:
   ```bash
   python migrate_script.py
   ```

## ผลลัพธ์

โปรแกรมจะสร้างไฟล์ `Match_Result_Final.xlsx` ที่มีข้อมูลดังนี้:

| คอลัมน์ | คำอธิบาย |
|---------|----------|
| `Oracle_ID` | รหัสลูกค้าจาก Oracle |
| `Oracle_Name` | ชื่อเต็มจาก Oracle |
| `Match_1_BP_Number` | รหัส BP ที่ใกล้เคียงที่สุด (อันดับ 1) |
| `Match_1_SAP_Name` | ชื่อที่ใกล้เคียงที่สุด (อันดับ 1) |
| `Match_1_Score` | คะแนนความเหมือน (0-100) |
| `Match_2_BP_Number` | รหัส BP อันดับ 2 |
| `Match_2_SAP_Name` | ชื่ออันดับ 2 |
| `Match_2_Score` | คะแนนอันดับ 2 |
| ... | (ไปจนถึงอันดับ 5) |

## การตีความคะแนน (Score)

- **90-100**: ใกล้เคียงมาก แนะนำให้ใช้
- **80-89**: ค่อนข้างใกล้เคียง ควรตรวจสอบ
- **70-79**: ใกล้เคียงปานกลาง ต้องตรวจสอบอย่างละเอียด
- **ต่ำกว่า 70**: อาจไม่ใช่คนเดียวกัน

## การแก้ไขปัญหา

### ไม่พบไฟล์ data_migration.xlsx
- ตรวจสอบว่าไฟล์อยู่ในโฟลเดอร์เดียวกับโปรแกรม
- ตรวจสอบชื่อไฟล์ว่าถูกต้อง

### Error เกี่ยวกับ Sheet
- ตรวจสอบว่ามี Sheet ชื่อ "Oracle" และ "SAP" ในไฟล์ Excel
- ตรวจสอบว่าชื่อ Sheet สะกดถูกต้อง

### Error เกี่ยวกับคอลัมน์
- ตรวจสอบว่ามีคอลัมน์ที่จำเป็นครบถ้วน:
  - Oracle Sheet: ID, Name1, Name2
  - SAP Sheet: BP_Number, Name1, Name2

## การปิด Virtual Environment

เมื่อใช้งานเสร็จแล้ว สามารถปิด Virtual Environment ได้ด้วยคำสั่ง:

```bash
deactivate
```

## ข้อมูลเพิ่มเติม

- โปรแกรมจะแสดงความคืบหน้าทุกๆ 50 รายการ
- ใช้เวลาประมาณ 1-5 วินาทีต่อรายการ (ขึ้นอยู่กับจำนวนข้อมูล)
- โปรแกรมจะข้ามคำที่ไม่จำเป็น เช่น บริษัท, จำกัด, คุณ, ฯลฯ เพื่อเพิ่มความแม่นยำ
# Fuzzword_check
