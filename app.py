from flask import Flask, request, Response
import pandas as pd
import os

EXCEL_FILE = "contacts.xlsx"
app = Flask(__name__)

# Load contacts
def load_excel():
    try:
        return pd.read_excel(EXCEL_FILE)
    except:
        return pd.DataFrame(columns=["company_name", "name", "mobile", "email"])

# Save contacts
def save_excel(df):
    df.to_excel(EXCEL_FILE, index=False)

# Insert new contact
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

    new_row = {
        "company_name": company,
        "name": name,
        "mobile": mobile,
        "email": email
    }
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    save_excel(df)
    return True

# Webhook endpoint
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
            search_term = message.split(" ", 1)[1].strip().lower()
            df = df.dropna(subset=["company_name"])  # إزالة الصفوف بدون اسم شركة
            df["company_name"] = df["company_name"].astype(str).str.lower().str.strip()

            results = df[df["company_name"].str.contains(search_term)]
            if results.empty:
                return twilio_reply("❌ لا توجد نتائج مطابقة.")

            reply = "\n".join([
                f"{row['name']} - {row['mobile']} - {row['email']}"
                for _, row in results.iterrows()
            ])
            return twilio_reply(f"📇 نتائج البحث:\n{reply}")

        except Exception as e:
            return twilio_reply(f"❌ تأكد من كتابة الأمر بالشكل: ابحث اسم_الشركة\n🔧 الخطأ: {str(e)}")


    elif "مساعدة" in message or "help" in message:
        return twilio_reply("🛠️ الأوامر المتاحة:\n- أضف الشركة, الاسم, الجوال, الإيميل\n- ابحث اسم_الشركة")

    else:
        return twilio_reply("❓ لم أفهم. أرسل 'مساعدة' لرؤية الأوامر المتاحة.")

# Function to return TwiML XML
def twilio_reply(message_text):
    xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<Response>
    <Message>{message_text}</Message>
</Response>"""
    return Response(xml, mimetype='application/xml')

# Start app
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
