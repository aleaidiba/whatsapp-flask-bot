from flask import Flask, request, jsonify
import pandas as pd
from difflib import SequenceMatcher
import os

EXCEL_FILE = "contacts.xlsx"

app = Flask(__name__)

def load_excel():
    try:
        return pd.read_excel(EXCEL_FILE)
    except:
        return pd.DataFrame(columns=["company_name", "name", "mobile", "email"])

def save_excel(df):
    df.to_excel(EXCEL_FILE, index=False)

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

@app.route("/", methods=["GET"])
def home():
    return "Contact Assistant Webhook is running."

@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json
    message = data.get("message", "").strip().lower()
    df = load_excel()

    if message.startswith("أضف "):
        try:
            _, content = message.split(" ", 1)
            parts = [x.strip() for x in content.split(",")]
            if len(parts) != 4:
                return jsonify({"reply": "❌ استخدم التنسيق: أضف الشركة, الاسم, الجوال, الإيميل"})
            company, name, mobile, email = parts
            added = insert_contact(df, company, name, mobile, email)
            return jsonify({"reply": "✅ تم الإضافة" if added else "⚠️ موجود مسبقاً"})
        except:
            return jsonify({"reply": "❌ حدث خطأ. تأكد من التنسيق: أضف الشركة, الاسم, الجوال, الإيميل"})

    elif message.startswith("ابحث "):
        company = message.replace("ابحث", "").strip().lower()
        results = df[df["company_name"].str.lower().str.contains(company)]
        if results.empty:
            return jsonify({"reply": "❌ لا توجد نتائج."})
        reply = "\n".join([f"{row['name']} - {row['mobile']} - {row['email']}" for _, row in results.iterrows()])
        return jsonify({"reply": f"📇 نتائج البحث:\n{reply}"})

    else:
        return jsonify({"reply": "أرسل 'أضف' لإضافة جهة أو 'ابحث' للبحث عن جهة."})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
