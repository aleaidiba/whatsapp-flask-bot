from flask import Flask, request, Response
import pandas as pd
import os
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

EXCEL_SHEET_NAME = "Contacts"  # اسم ملف Google Sheet

app = Flask(__name__)

# الاتصال بـ Google Sheets باستخدام GOOGLE_CREDENTIALS من المتغير البيئي
def connect_to_sheet():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds_json = os.getenv("GOOGLE_CREDENTIALS")
    creds_dict = json.loads(creds_json)
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    sheet = client.open(EXCEL_SHEET_NAME).sheet1
    return sheet

# تحميل البيانات
def load_excel():
    sheet = connect_to_sheet()
    records = sheet.get_all_records()
    return pd.DataFrame(records)

# إدخال جهة اتصال جديدة
def insert_contact(df, company, name, mobile, email):
    try:
        sheet = connect_to_sheet()
        mobile = str(mobile or "")
        email = str(email or "")

        duplicate = df[
            (df["name"].str.lower() == name.lower()) |
            (df["email"].str.lower() == email.lower()) |
            (df["mobile"].astype(str) == mobile)
        ]
        if not duplicate.empty:
            print("⚠️ جهة الاتصال مكررة")
            return False

        sheet.append_row([company, name, mobile, email])
        print("✅ تم إدخال جهة الاتصال في Google Sheets")
        return True
    except Exception as e:
        print(f"❌ خطأ أثناء الإضافة: {e}")
        return False

# رد Twilio بصيغة XML
def twilio_reply(message_text):
    xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<Response>
    <Message>{message_text}</Message>
</Response>"""
    return Response(xml, mimetype='application/xml')

# نقطة الدخول Webhook
@app.route("/webhook", methods=["POST"])
def webhook():
    message = request.form.get("Body", "").strip().lower()
    df = load_excel()

    if message.startswith("أضف "):
        try:
            _, content = message.split(" ", 1)
            parts = [x.strip() for x in content.split(",")]
            if len(parts) != 4:
                return twilio_reply("❌ استخدم التنسيق: أضف الشركة, الاسم, الجوال, الإيميل")
            company, name, mobile, email = parts
            added = insert_contact(df, company, name, mobile, email)
            return twilio_reply("✅ تم الإضافة" if added else "⚠️ موجود مسبقاً")
        except Exception as e:
            return twilio_reply(f"❌ خطأ أثناء الإضافة: {str(e)}")

    elif message.startswith("ابحث"):
        try:
            parts = message.split(" ", 1)
            if len(parts) < 2 or not parts[1].strip():
                return twilio_reply("❌ اكتب اسم الشركة بعد كلمة 'ابحث'. مثل: ابحث شركة الاختبار")

            search_term = parts[1].strip().lower()
            df["company_name"] = df["company_name"].fillna('').astype(str).str.lower().str.strip()
            results = df[df["company_name"].str.contains(search_term)]

            if results.empty:
                return twilio_reply("❌ لا توجد نتائج مطابقة.")

            reply = "\n".join([
                f"{row['name']} - {row['mobile']} - {row['email']}"
                for _, row in results.iterrows()
            ])
            return twilio_reply(f"📇 نتائج البحث:\n{reply}")

        except Exception as e:
            return twilio_reply(f"❌ خطأ أثناء البحث: {str(e)}")

    elif "مساعدة" in message or "help" in message:
        return twilio_reply("🛠️ الأوامر المتاحة:\n- أضف الشركة, الاسم, الجوال, الإيميل\n- ابحث اسم_الشركة")

    else:
        return twilio_reply("❓ لم أفهم. أرسل 'مساعدة' لرؤية الأوامر المتاحة.")

# صفحة رئيسية للفحص
@app.route("/", methods=["GET"])
def home():
    return "✅ WhatsApp Flask Bot + Google Sheets يعمل باستخدام GOOGLE_CREDENTIALS"

# تشغيل التطبيق
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
