pip install fastapi uvicorn python-multipart


import pandas as pd

import streamlit as st
import openai
import json
import re
import extract_msg
import os
from pathlib import Path
from difflib import SequenceMatcher

openai.api_key = os.getenv("OPENAI_API_KEY")
EXCEL_FILE = "contacts.xlsx"

# --- Load and Save Excel ---
def load_excel():
    try:
        return pd.read_excel(EXCEL_FILE, engine='openpyxl')
    except FileNotFoundError:
        df = pd.DataFrame(columns=["company_name", "name", "mobile", "email"])
        df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')
        return df
    except Exception as e:
        st.error(f"❌ Failed to load Excel file: {e}")
        return pd.DataFrame(columns=["company_name", "name", "mobile", "email"])


def save_excel(df):
    df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')
# --- Insert Contact with Duplicate Check ---
# --- Insert Contact with Duplicate Check ---
def insert_contact(df, company, name, mobile, email):
    # Ensure all values are not None and are strings
    company = company or ""
    name = name or ""
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


# --- Extract Contact from .msg File ---
def extract_contact_from_msg(msg_file_path):
    msg = extract_msg.Message(msg_file_path)
    body = msg.body

    gpt_prompt = f"""
You are an Outlook signature extraction expert working for Solutions by STC.

From the following email body, extract ONLY the contact details of people who are not employees of solutions.com.sa:

Email:
{body}

Only extract one person's contact from the signature (not the body of the message), and return it in this JSON format:
{{
  "company_name": "<company>",
  "name": "<full name>",
  "mobile": "<preferably Saudi mobile>",
  "email": "<email>"
}}

If no valid external contact is found, return an empty JSON like: {{}}
"""
    try:
        response = openai.chat.completions.create(
            model="gpt-4o",
            temperature=0,
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": gpt_prompt}
            ]
        )
        result = response.choices[0].message.content
        match = re.search(r'\{.*\}', result, re.DOTALL)
        if match:
            contact = json.loads(match.group())
            if contact.get("email") and not contact["email"].endswith("@solutions.com.sa"):
                return contact
    except Exception as e:
        print(f"Error in GPT extraction: {e}")
    return None
from fastapi import FastAPI, Request, Form
from fastapi.responses import JSONResponse

app = FastAPI()


@app.post("/whatsapp/webhook")
async def whatsapp_webhook(
    request: Request,
    Body: str = Form(default=""),
    From: str = Form(default=""),
    WaId: str = Form(default=""),
    ProfileName: str = Form(default=""),
    ButtonResponse: str = Form(default=""),  # When user responds to a button
):
    print("📨 Received message:")
    form_data = await request.form()
    for key, value in form_data.items():
        print(f"{key}: {value}")

    # Optionally extract the interactive reply
    payload = form_data.get("ListResponse")
    if payload:
        selected_id = form_data.get("ListResponse.Id")
        print(f"✅ User selected option: {selected_id}")

        if selected_id == "chat_assistant":
            return JSONResponse({"reply": "🧠 مرحبًا بك في مساعد جهات الاتصال. كيف أستطيع مساعدتك؟"})
        elif selected_id == "view_contacts":
            return JSONResponse({"reply": "👁️ هذه قائمة جهات الاتصال المحفوظة لدينا..."})

    # Default fallback
    return JSONResponse({"reply": "👋 مرحبًا بك. أرسل 'قائمة' لعرض الخيارات."})

# --- GPT Assistant for Chat ---
def process_with_llm(user_input):
    system_message = """
You are a helpful assistant that manages a contact database stored in Excel.
Each contact has: company_name, name, mobile, email.

You can perform these actions:
1. Retrieve — Get all contacts for a company.
2. Insert — Add a new contact.
3. Update — Modify mobile or email for an existing contact.
4. Delete — Remove a contact by name and company.
5. Update a cell — Change a specific value by row index and column.

Return only a valid JSON object. Do not include any explanation or extra text.
Examples:
{"action": "retrieve", "company": "Cisco"}
{"action": "insert", "company": "STC", "name": "Ali", "mobile": "0555000000", "email": "ali@stc.com"}
"""
    try:
        response = openai.chat.completions.create(
            model="gpt-4o",
            temperature=0,
            messages=[
                {"role": "system", "content": system_message},
                {"role": "user", "content": user_input}
            ]
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return json.dumps({"error": f"OpenAI API error: {str(e)}"})

# --- Main Menu ---
def show_main_menu():
    return st.radio("Choose a task:", [
        "Upload Excel with company names",
        "Upload full contact Excel",
        "Upload single .msg file",
        "Upload folder of .msg files",
        "Extract from BCK File",
        "Extract from WhatsApp Group",
        "Chat Assistant",
        "View Contacts"
    ])

# --- GPT Fuzzy Matching for Company Names ---
def get_gpt_similar_names(input_name, existing_names):
    prompt = f"""
You are a fuzzy matching assistant.
From the list below, return all names that closely or partially match the term: '{input_name}'
Avoid returning names that are similar to '{input_name}' only because they share common suffixes or prefixes.

List:
{existing_names}

Return the result as a JSON array of strings.
"""
    try:
        response = openai.chat.completions.create(
            model="gpt-4o",
            temperature=0,
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": prompt}
            ]
        )
        matches = json.loads(re.search(r'\[.*\]', response.choices[0].message.content, re.DOTALL).group())
        return matches
    except:
        return []

# --- Fuzzy Matching Heuristic ---
def is_valid_match(query, candidate):
    query, candidate = query.lower().strip(), candidate.lower().strip()
    if query in candidate or candidate in query:
        return True
    if candidate.startswith(query.split()[0]) or query.startswith(candidate.split()[0]):
        return True
    ratio = SequenceMatcher(None, query, candidate).ratio()
    return ratio > 0.75

# --- Agent 1: Auto-fill from Database ---
def agent_autofill_company_names(df):
    st.subheader("📄 Upload Excel to Auto-fill from Local Database")
    uploaded_file = st.file_uploader("Upload Excel with 'company_name' column", type=["xlsx"], key="autofill")
    if uploaded_file:
        uploaded_df = pd.read_excel(uploaded_file)
        if "company_name" not in uploaded_df.columns:
            st.error("❌ Missing 'company_name' column in uploaded file.")
            return

        db = load_excel()
        db["company_name"] = db["company_name"].astype(str).str.lower().str.strip()
        autofilled_rows = []
        existing_names = db["company_name"].dropna().unique().tolist()

        for _, input_row in uploaded_df.iterrows():
            company = str(input_row["company_name"]).strip()
            if not company:
                continue

            matches = db[db["company_name"] == company.lower()]

            if matches.empty:
                gpt_matches = get_gpt_similar_names(company, existing_names)
                filtered = [m for m in gpt_matches if is_valid_match(company, m)]
                matches = db[db["company_name"].isin([m.lower().strip() for m in filtered])]

            if not matches.empty:
                for _, match_row in matches.iterrows():
                    autofilled_rows.append({
                        "company_name": company,
                        "matched_company": match_row["company_name"],
                        "name": match_row["name"],
                        "mobile": match_row["mobile"],
                        "email": match_row["email"]
                    })
            else:
                autofilled_rows.append({
                    "company_name": company,
                    "matched_company": "",
                    "name": "",
                    "mobile": "",
                    "email": ""
                })

        if autofilled_rows:
            result_df = pd.DataFrame(autofilled_rows)
            st.success(f"✅ Auto-filled {len(result_df)} record(s).")
            st.dataframe(result_df)
            output_file = "autofilled_contacts_from_db.xlsx"
            result_df.to_excel(output_file, index=False)
            with open(output_file, "rb") as f:
                st.download_button("📥 Download Updated Excel", f, file_name=output_file)

# --- Agent 2: Upload Full Contact Excel ---
def agent_upload_full_excel(df):
    st.subheader("📤 Upload Full Contact Excel")
    file = st.file_uploader("Upload Excel with full info", type=["xlsx"], key="full")
    if file:
        try:
            uploaded_df = pd.read_excel(file)
            required = {"company_name", "name", "mobile", "email"}
            if not required.issubset(uploaded_df.columns):
                st.error("❌ Must include: company_name, name, mobile, email")
                return

            added, updated, skipped = 0, 0, 0
            for _, row in uploaded_df.iterrows():
                email = row["email"]
                match = df[df["email"].str.lower() == str(email).lower()]
                if match.empty:
                    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
                    added += 1
                else:
                    idx = match.index[0]
                    if not row.equals(match.iloc[0]):
                        df.loc[idx] = row
                        updated += 1
                    else:
                        skipped += 1
            save_excel(df)
            st.success(f"✅ Added: {added}, Updated: {updated}, Skipped: {skipped}")
        except Exception as e:
            st.error(f"❌ Error reading file: {e}")

# --- Agent 3: Upload single .msg file ---
def agent_upload_msg_file(df):
    st.subheader("📨 Upload Single .msg File")
    file = st.file_uploader("Upload .msg file", type=["msg"], key="msg")
    if file:
        with open("temp.msg", "wb") as f:
            f.write(file.read())
        contact = extract_contact_from_msg("temp.msg")
        if contact:
            st.success("✅ Contact found. Confirm or edit:")
            edited_df = st.data_editor(pd.DataFrame([contact]), num_rows="dynamic")
            if st.button("💾 Save Contact"):
                row = edited_df.iloc[0]
                if insert_contact(df, row["company_name"], row["name"], row["mobile"], row["email"]):
                    st.success("✅ Contact saved.")
                else:
                    st.info("ℹ️ Duplicate contact skipped.")

# --- Agent 4: WhatsApp Export (.txt) ---
def agent_extract_whatsapp_bundle(df):
    st.subheader("📱 Upload WhatsApp Group Export File")
    uploaded_file = st.file_uploader("📤 Drag and drop your WhatsApp group chat (.txt) file here", type=["txt"], key="whatsapp_group")

    extracted = []

    def process_chunk(chunk_text):
        gpt_prompt = f"""
You are an assistant that extracts contact details from a WhatsApp group text export.
Extract available fields: company_name, name, mobile, and email from all visible contact card blocks.
Return a JSON list of contacts like:
[
  {{"company_name": "<optional>", "name": "<required>", "mobile": "<required>", "email": "<optional>"}},
  ...
]

Text:
{chunk_text}
"""
        try:
            response = openai.chat.completions.create(
                model="gpt-4o",
                temperature=0,
                messages=[
                    {"role": "system", "content": "You are a helpful assistant."},
                    {"role": "user", "content": gpt_prompt}
                ]
            )
            return json.loads(re.search(r'\[.*\]', response.choices[0].message.content, re.DOTALL).group())
        except Exception as e:
            st.error(f"❌ Error in chunk: {e}")
            return []

    if uploaded_file:
        try:
            content = uploaded_file.read().decode("utf-8", errors="ignore")
            lines = content.splitlines()
            chunk_size = 300  # ~300 lines per chunk
            chunks = [lines[i:i + chunk_size] for i in range(0, len(lines), chunk_size)]

            st.info(f"⏳ Processing {len(chunks)} chunks...")

            for i, chunk in enumerate(chunks):
                chunk_text = "\n".join(chunk)
                contacts = process_chunk(chunk_text)
                for contact in contacts:
                    if contact.get("name") and contact.get("mobile"):
                        extracted.append(contact)
                st.write(f"✅ Processed chunk {i+1} of {len(chunks)}")

        except Exception as e:
            st.error(f"❌ Failed to extract from WhatsApp group: {e}")

    if extracted:
        df_new = pd.DataFrame(extracted)
        st.success(f"✅ Extracted {len(df_new)} contact(s) from WhatsApp group.")
        st.dataframe(df_new)

        # 🔁 Automatically insert into main DB
        saved_count = 0
        for _, row in df_new.iterrows():
            if insert_contact(df, row.get("company_name", ""), row["name"], row["mobile"], row.get("email", "")):
                saved_count += 1

        # 💾 Save new contacts to temp Excel file
        output_file = "whatsapp_contacts_extracted.xlsx"
        df_new.to_excel(output_file, index=False)

        # 🎉 Show success and download
        st.success(f"💾 Automatically saved {saved_count} contact(s) to database.")
        with open(output_file, "rb") as f:
            st.download_button(
                label="📥 Download Extracted Contacts Excel",
                data=f,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("⚠️ No valid contacts found.")

# --- Main ---
def main():
    st.set_page_config(layout="wide")
    st.title("🤖 Multi-Agent Contact Assistant")
    df = load_excel()
    task = show_main_menu()

    if task == "Upload Excel with company names":
        agent_autofill_company_names(df)
    elif task == "Upload full contact Excel":
        agent_upload_full_excel(df)
    elif task == "Upload single .msg file":
        agent_upload_msg_file(df)
    elif task == "Extract from WhatsApp Group":
        agent_extract_whatsapp_bundle(df)
    elif task == "Chat Assistant":
        agent_chat_assistant(df)
    elif task == "View Contacts":
        st.dataframe(df)

if __name__ == "__main__":
    main()
