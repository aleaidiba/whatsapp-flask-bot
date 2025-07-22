from flask import Flask, request, Response
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os

app = Flask(__name__)

# إعداد Google Sheets
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
CREDS_JSON = os.environ.get("GOOGLE_CREDENTIALS")

# تحويل JSON من string إلى ملف مؤقت
import json, tempfile
with tempfile.NamedTemporaryFile(delete=False, mode='w+', suffix='.json') as tmp:
    tmp.write(CREDS_JSON)
    CREDENTIALS_FILE = tmp.name

SPREADSHEET_NAME = "contacts"  # تأكد أن هذا الاسم يطابق اسم Google Sheet

def connect_to_sheet():
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, SCOPE)
    client = gspread.authorize(creds)
    sheet = client.open(SPREADSHEET_NAME).sheet1
    return sheet

def load_excel():
    sheet = connect_to_sheet()
    records = sheet.get_all_records()
    return pd.DataFrame(records)

def save_to_sheet(company, name, mobile, email):
    sheet = connect_to_sheet()
    sheet.append_row([company, name, mobile, email])

def insert_contact(df, company, name, mobile, email):
    mobile = str(mobile or "")
    email = str(email or "")

    duplicate = df[
        (df["name"].str.lower() == name.lower()) |
        (df["email"].str.lower() == email.lower()) |
        (df["mobile"].astype(str) == mobile)
    ]
    if not duplicate.empty:
        return False

    save_to_sheet(company, name, mobile, email)
    return True

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
        except:
            return twilio_reply("❌ حدث خطأ. تأكد من التنسيق: أضف الشركة, الاسم, الجوال, الإيميل")

    elif message.startswith("ابحث "):
        try:
            company = message.replace("ابحث", "").strip().lower()
            df.dropna(subset=["company_name"], inplace=True)
            results = df[df["company_name"].str.lower().str.contains(company)]
            if results.empty:
                return twilio_reply("❌ لا توجد نتائج.")
            reply = "\n".join([f"{row['name']} - {row['mobile']} - {row['email']}" for _, row in results.iterrows()])
            return twilio_reply(f"🗂️ نتائج البحث:\n{reply}")
        except Exception as e:
            return twilio_reply(f"⚠️ خطأ في البحث: {e}")

    elif "مساعدة" in message or "help" in message:
        return twilio_reply("🛠️ الأوامر المتاحة:\n- أضف الشركة, الاسم, الجوال, الإيميل\n- ابحث اسم_الشركة")

    else:
        return twilio_reply("❓ لم أفهم. أرسل 'مساعدة' لرؤية الأوامر المتاحة.")

def twilio_reply(message_text):
    xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<Response>
    <Message>{message_text}</Message>
</Response>"""
    return Response(xml, mimetype='application/xml')

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
