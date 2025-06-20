import os
import requests
import json
from dotenv import load_dotenv
import base64
from datetime import datetime, timedelta
import time
import csv
import pandas as pd

# Load environment variables from .env
load_dotenv()

def create_ticket_from_exact_data(thai_datetime_str, system_status_description):
    """
    Create Freshservice ticket with Thai datetime input and description.
    """

    # ตัวอย่าง .env:
    # FRESHSERVICE_DOMAIN = "your_domain.freshservice.com"
    # FRESHSERVICE_API_KEY = "your_api_key_here"

    domain = os.getenv('FRESHSERVICE_DOMAIN')  # e.g., "yourcompany.freshservice.com"
    api_key = os.getenv('FRESHSERVICE_API_KEY')  # e.g., "abc123apikey"

    credentials = base64.b64encode(f"{api_key}:X".encode()).decode()

    headers = {
        'Content-Type': 'application/json; charset=utf-8',
        'Authorization': f'Basic {credentials}'
    }

    def thai_date_to_utc_iso(thai_date_str):
        try:
            thai_date_str = thai_date_str.replace(" ", ":")
            date_part, time_part = thai_date_str.split(':', 1)
            year, month, day = map(int, date_part.split('-'))
            hour, minute, *sec = map(int, time_part.split(':'))
            second = sec[0] if sec else 0

            gregorian_year = year if year < 2500 else year - 543
            local_dt = datetime(gregorian_year, month, day, hour, minute, second)
            utc_dt = local_dt - timedelta(hours=7)

            return utc_dt.strftime('%Y-%m-%dT%H:%M:%S.000Z')
        except Exception as e:
            print("❌ Error parsing Thai date:", e)
            return datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%S.000Z')

    current_time = thai_date_to_utc_iso(thai_datetime_str)

    ticket_data = {
        "subject": "รายงานสถานะระบบเครือข่ายและเว็บไซต์",
        "description": system_status_description,
        "email": "example@domain.com",  # 🔒 แทนอีเมลจริง
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
            "cf_internal": "ชื่อผู้รับผิดชอบ",
            "category_1": "กลุ่มงานหลัก",
            "category_2": "งานที่เกี่ยวข้อง",
            "rand37897016": current_time,  # 🔒 วันที่ในรูปแบบ ISO
            "rand13178744": "ค่าเริ่มต้น",
            "rand30247461": "รายละเอียดงาน",
            "rand30541414": "รายละเอียดเพิ่มเติม",
            "rand45585306": "แหล่งที่มา",
            "rand19697592": "ตำแหน่ง",
            "rand35121958": "ชื่อผู้รายงาน",
            "ticket_type": "Daily Job",
            "rand30625773": "ประเภทผู้ใช้",
            "rand55224974": "หมวดหมู่เพิ่มเติม",
        }
    }

    url = f"https://{domain}/api/v2/tickets"
    try:
        response = requests.post(
            url,
            headers=headers,
            data=json.dumps(ticket_data, ensure_ascii=False).encode('utf-8')
        )

        if response.status_code == 201:
            result = response.json()
            print("✅ สร้างตั๋วสำเร็จ!")
            print(f"🎫 Ticket ID: {result['ticket']['id']}")
            print(f"🔗 URL: https://{domain}/a/tickets/{result['ticket']['id']}")
            return result
        else:
            print(f"❌ Error: {response.status_code}")
            print(f"📝 Response: {response.text}")
            return None

    except Exception as e:
        print(f"❌ Connection Error: {str(e)}")
        return None

def update_data(excel_file, system_status_description, old_value="Accomplished", new_value="OK", delay=5):
    df = pd.read_excel(excel_file, engine='openpyxl')
    updated_rows = 0
    for i in range(len(df)):
        if df.at[i, "_complete"] == old_value:
            df.at[i, "_complete"] = new_value
            updated_rows += 1
            print(f"✅ อัปเดต _complete ที่แถว {i} (วันที่: {df.at[i, '_date']})")
            time.sleep(delay)
            df.to_excel(excel_file, index=False, engine='openpyxl')
            create_ticket_from_exact_data(df.at[i, '_date'], system_status_description)
            time.sleep(delay)
    print(f"รวมอัปเดตทั้งหมด: {updated_rows} แถว")

if __name__ == "__main__":
    print("🎯 เริ่มสร้างตั๋ว Freshservice สำหรับเมืองพัทยา")
    system_status_description = "<html>...ข้อมูลสถานะระบบที่ใช้ได้ทั้งหมด เช่น IP Phone, DNS, Website ฯลฯ...</html>"  # 🔒 แทน HTML ยาว
    excel_file = "working_days.xlsx"
    update_data(excel_file, system_status_description, old_value="Accomplished", new_value="OK", delay=1)
