from flask import Flask, request, jsonify
import requests
import smtplib
import json
import openai
import os
from email.mime.text import MIMEText
from msal import ConfidentialClientApplication
from telegram import Bot
from telegram.ext import Updater, CommandHandler, MessageHandler, Filters
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Configuration
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

app = Flask(__name__)

# Set up OpenAI API
openai.api_key = OPENAI_API_KEY

# Set up Telegram Bot
bot = Bot(token=TELEGRAM_BOT_TOKEN)

# Microsoft Graph API Authentication
def get_access_token():
    app = ConfidentialClientApplication(CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}",
                                        client_credential=CLIENT_SECRET)
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return token.get("access_token")

# Fetch unread emails
def fetch_unread_emails():
    access_token = get_access_token()
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$filter=isRead eq false"

    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        emails = response.json().get("value", [])
        return emails
    else:
        return []

# Generate AI Response
def generate_ai_response(user_query):
    prompt = f"Client Issue: {user_query}\nProvide a professional and helpful support response."
    
    response = openai.ChatCompletion.create(
        model="gpt-4",
        messages=[{"role": "system", "content": prompt}]
    )

    return response['choices'][0]['message']['content']

# Send email response
def send_email(to_email, subject, body):
    smtp_server = "smtp.office365.com"
    smtp_port = 587
    smtp_user = EMAIL_ADDRESS
    smtp_password = "your-app-password"

    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = smtp_user
    msg["To"] = to_email

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_user, to_email, msg.as_string())

# Telegram Bot Handlers
def start(update, context):
    update.message.reply_text("Hello! I will notify you when a new support email arrives.")

def handle_message(update, context):
    text = update.message.text
    chat_id = update.message.chat_id
    
    if text.startswith("APPROVE:"):
        data = text.replace("APPROVE:", "").strip().split("|")
        to_email, subject, body = data
        send_email(to_email, subject, body)
        update.message.reply_text("✅ Email sent successfully.")
    
    elif text.startswith("EDIT:"):
        data = text.replace("EDIT:", "").strip().split("|")
        to_email, subject, new_body = data
        send_email(to_email, subject, new_body)
        update.message.reply_text("✅ Edited response sent successfully.")

def send_telegram_notification(email_data, ai_response):
    sender = email_data["from"]["emailAddress"]["address"]
    subject = email_data["subject"]
    message_body = email_data["body"]["content"]

    message = (
        f"📩 New Support Request\n"
        f"👤 From: {sender}\n"
        f"📝 Subject: {subject}\n"
        f"📖 Message: {message_body}\n\n"
        f"🤖 AI Suggested Response:\n"
        f"{ai_response}\n\n"
        f"Reply with:\n"
        f"✅ APPROVE:{sender}|{subject}|{ai_response}\n"
        f"✏️ EDIT:{sender}|{subject}|[Your Edited Response]"
    )

    bot.send_message(chat_id=TELEGRAM_CHAT_ID, text=message)

# Process emails and send notifications
@app.route('/process-emails', methods=['GET'])
def process_emails():
    emails = fetch_unread_emails()
    if not emails:
        return jsonify({"message": "No unread emails."})

    for email in emails:
        sender = email["from"]["emailAddress"]["address"]
        subject = email["subject"]
        body = email["body"]["content"]

        ai_response = generate_ai_response(body)

        send_telegram_notification(email, ai_response)

    return jsonify({"message": "Telegram notifications sent for approval."})

# Start Telegram Bot
def start_telegram_bot():
    updater = Updater(TELEGRAM_BOT_TOKEN, use_context=True)
    dp = updater.dispatcher

    dp.add_handler(CommandHandler("start", start))
    dp.add_handler(MessageHandler(Filters.text & ~Filters.command, handle_message))

    updater.start_polling()
    updater.idle()

# Run the Flask API
if __name__ == '__main__':
    from threading import Thread

    telegram_thread = Thread(target=start_telegram_bot)
    telegram_thread.start()

    app.run(port=5000, debug=True)
