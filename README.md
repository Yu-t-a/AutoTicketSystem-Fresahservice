# 📄 README.md - Freshservice Ticket Automation

## 📌 รายละเอียดระบบ

สคริปต์นี้ใช้สำหรับการสร้าง Ticket โดยอัตโนมัติบน Freshservice ผ่าน API
โดยสามารถระบุวันที่ (ในรูปแบบ พ.ศ.) และรายละเอียดสถานะระบบที่ต้องการแนบเข้าไปยังช่อง Description ได้โดยตรง

---

## ⚙️ ความต้องการของระบบ

* Python 3.8+
* แพ็กเกจ Python:

  * requests
  * python-dotenv
  * pandas
  * openpyxl
* ไฟล์ `.env` ต้องประกอบด้วย:

  ```env
  FRESHSERVICE_DOMAIN=yourcompany.freshservice.com
  FRESHSERVICE_API_KEY=your_api_key_here
  ```

---

## 🧾 ตัวอย่างการใช้งาน

ให้เตรียมไฟล์ Excel ชื่อ `working_days.xlsx` ที่มีคอลัมน์ `_date`, `_worker`, `_complete`

## คำอธิบายรูปแบบไฟล์ working_days.xlsx
```
| คอลัมน์     | ประเภทข้อมูล                 | ตัวอย่าง                 | คำอธิบาย                                                                                                     |
| ----------- | ---------------------------- | ------------------------ | ------------------------------------------------------------------------------------------------------------ |
| `_date`     | วันที่-เวลา (str)            | `2568-06-19 09:00`       | วันที่และเวลาที่จะใช้เป็นวันที่ในตั๋ว (รูปแบบ **พ.ศ.** ที่ต้องแปลงเป็น ค.ศ.)                                 |
| `_worker`   | รายละเอียดงาน | `รายละเอียดงาน`   | (ไม่บังคับใช้ในสคริปต์ปัจจุบัน แต่สามารถใช้เพื่อบันทึกชื่อผู้รับผิดชอบในอนาคตได้)                            |
| `_complete` | สถานะการดำเนินการ (str)      | `progress` หรือ `OK` | ถ้าเป็น `progress` หมายถึง **ยังไม่ได้สร้างตั๋ว** → ระบบจะสร้างตั๋วและอัปเดตเป็น `OK` ทันทีหลังส่งสำเร็จ |

```

```bash
python ticket_auto.py
```

ระบบจะอัปเดตสถานะจาก `progress` → `OK` และสร้างตั๋วใหม่ใน Freshservice

---

## 🧠 ฟังก์ชันสำคัญ

### `create_ticket_from_exact_data(thai_datetime_str, system_status_description)`

* ใช้สำหรับสร้างตั๋ว โดยแปลงวันที่จากรูปแบบ พ.ศ. เช่น `2568-06-19 09:00` → UTC ISO format
* แนบคำอธิบาย HTML สำหรับแสดงสถานะระบบ

### ตัวอย่าง `ticket_data` ที่ใช้จริง:

```json
{
  "subject": "รายงานสถานะระบบเครือข่ายและเว็บไซต์",
  "description": "<div>...HTML รายงาน...</div>",
  "email": "example@domain.com",
  "priority": 1,
  "status": 5,
  "source": 1002,
  "urgency": 1,
  "impact": 1,
  "category": "Service Support IT",
  "sub_category": "Website",
  "item_category": "Other",
  "type": "Service Request",
  "workspace_id": 2,
  "custom_fields": {
    "cf_internal": "นาย[...] [...]",
    "category_1": "งานบริหารจัดการและตรวจสอบระบบ",
    "category_2": "เว็บไซต์และแอปพลิเคชั่น",
    "rand37897016": "2025-06-19T02:00:00.000Z",
    "rand13178744": "Default Value",
    "rand30247461": "งานตรวจสอบ Server และ Network",
    "rand30541414": "ตรวจสอบ Server Website",
    "rand45585306": "CCR-Internal",
    "rand19697592": "เว็บโปรแกรมเมอร์",
    "rand35121958": "นาย [...] [...]",
    "ticket_type": "Daily Job",
    "rand30625773": "พนักงาน",
    "rand55224974": "อื่น ๆ"
  }
}
```

---

## 📂 โครงสร้างไฟล์

```
project/
├── ticket_auto.py
├── working_days.xlsx
├── .env
└── README.md
```

---

## 🛠️ ข้อแนะนำเพิ่มเติม

* ให้ระบุฟิลด์ custom ให้ครบทุกตัวที่ระบบ Freshservice กำหนด (ใช้ create\_ticket\_with\_debug เพื่อตรวจสอบฟิลด์ที่ขาด)
* กำหนด delay เพื่อไม่ให้ยิง API ถี่เกินไป (default = 1-5 วินาที)
* ข้อความใน `description` ควรเป็น HTML ที่จัดรูปแบบให้พร้อมแสดงบนหน้า Freshservice

---

## 🧪 ทดสอบระบบด้วยฟังก์ชัน

```python
create_ticket_with_debug()
```

หาก API แจ้งฟิลด์ไม่ครบหรือผิดประเภท จะทำการ Debug โดยแสดงชื่อฟิลด์และข้อความผิดพลาด

---

## 🧑‍💻 ผู้พัฒนา

* ผู้พัฒนา: Yu (Weerayut J.)
