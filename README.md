# DLT Statistics Collector

โปรแกรม Terminal UI สำหรับดาวน์โหลด Excel สถิติด้านทะเบียนรถจาก [กรมการขนส่งทางบก](https://web.dlt.go.th/statistics/) และจัดเก็บแยกตามหมวด / หัวข้อ / ค่า Dropdown

## วิธีใช้งาน (PowerShell)

**1. Clone และเข้าโฟลเดอร์**

```powershell
git clone https://github.com/firmeen/DLT.git
cd DLT
```

**2. ติดตั้งแพ็กเกจ**

```powershell
python -m pip install -r requirements.txt
```

**3. เปิดโปรแกรม**

```powershell
python dlt_collector.py
```

**4. ใช้งานใน Terminal UI**

- กด **DISCOVER** เพื่ออ่านรายการจากเว็บ
- เลือกหมวด แล้วกด **DOWNLOAD ALL** หรือเลือกแถวแล้วกด **DOWNLOAD ROW**
- กด **STOP** เพื่อหยุด; **Ctrl+Q** เพื่อออก

ไฟล์อยู่ในโฟลเดอร์ `DLT_Statistics/` พร้อม `manifest.csv` รันซ้ำจะข้ามไฟล์ที่ดาวน์โหลดแล้วอัตโนมัติ

> หากพบ SSL error ให้ตรวจสอบ certificate ของ Windows/เครือข่ายก่อน ไม่ควรปิด SSL verification ยกเว้นทดสอบชั่วคราว
