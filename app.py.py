#Test
from flask import Flask, request
from twilio.twiml.messaging_response import MessagingResponse

app = Flask(__name__)

@app.route("/whatsapp", methods=["POST"])
def whatsapp_reply():
    incoming_msg = request.values.get("Body", "")
    resp = MessagingResponse()
    msg = resp.message()
    msg.body(f"You said: {incoming_msg}")
    return str(resp)
