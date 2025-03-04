import os
import requests
import asyncio
import re
import html
import threading
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from msal import ConfidentialClientApplication
from flask import Flask, jsonify
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes

# 🔹 Load environment variables
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")  
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

# 🔹 Microsoft Graph API Endpoints
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
EMAILS_URL = f"https://graph.microsoft.com/v1.0/users/{EMAIL_ADDRESS}/messages"

# 🔹 Email Storage & AI Memory
email_store = {}
ai_responses = {}  # Stores past AI-generated responses for improvement

# 🔹 Flask App Setup
flask_app = Flask(__name__)

@flask_app.route("/")
def home():
    return jsonify({"message": "Support Bot is running!"})

# 🔹 Get Access Token
def get_access_token():
    """Fetch access token using client credentials"""
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }

    response = requests.post(TOKEN_URL, data=data)

    if response.status_code == 200:
        return response.json().get("access_token")
    else:
        return None

# 🔹 Fetch Unread Emails (With Time Filter)
def fetch_unread_emails(access_token, hours=None):
    """Fetch unread emails from Microsoft Graph API, filtering out sent emails and limiting time range."""
    headers = {"Authorization": f"Bearer {access_token}"}
    
    query = "isRead eq false"
    if hours:
        time_filter = (datetime.utcnow() - timedelta(hours=int(hours))).isoformat() + "Z"
        query += f" and receivedDateTime ge {time_filter}"

    response = requests.get(
        f"{EMAILS_URL}?$filter={query}&$orderby=receivedDateTime desc",
        headers=headers
    )

    if response.status_code == 200:
        emails = response.json().get("value", [])
        filtered_emails = [email for email in emails if email.get("from", {}).get("emailAddress", {}).get("address") != EMAIL_ADDRESS]
        return filtered_emails
    else:
        return None

# 🔹 Send Messages to Telegram (Using HTML Mode)
def send_to_telegram(message):
    """Send a well-formatted message to Telegram using HTML mode."""
    telegram_url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"

    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": html.escape(message),
        "parse_mode": "HTML",
        "disable_web_page_preview": True
    }

    response = requests.post(telegram_url, json=payload)
    return response.status_code == 200

# 🔹 Async Function for Email Processing
async def send_email_to_telegram(hours=None):
    """Fetch unread emails and send them to Telegram."""
    access_token = get_access_token()
    if not access_token:
        send_to_telegram("❌ Could not retrieve access token.")
        return

    emails = fetch_unread_emails(access_token, hours)
    if not emails:
        send_to_telegram("📭 No new unread emails found.")
        return

    email_store.clear()
    for index, email in enumerate(emails[:10], start=1):  # Shows up to 10 emails
        subject = email.get("subject", "No Subject")
        sender_name = email.get("from", {}).get("emailAddress", {}).get("name", "Unknown Sender")
        sender_email = email.get("from", {}).get("emailAddress", {}).get("address", "Unknown Email")
        received_time = email.get("receivedDateTime", "Unknown Time")
        body_html = email.get("body", {}).get("content", "No Preview Available")

        soup = BeautifulSoup(body_html, "html.parser")
        body_text = soup.get_text()

        email_store[str(index)] = {"sender": sender_name, "subject": subject, "body": body_text}

        message = (
            f"<b>📩 New Email Received</b> [#{index}]\n"
            f"📌 <b>From:</b> {sender_name} ({sender_email})\n"
            f"📌 <b>Subject:</b> {subject}\n"
            f"🕒 <b>Received:</b> {received_time}\n"
            f"📝 <b>Preview:</b> {body_text[:500]}...\n\n"
            f"✍️ Reply with: <code>/suggest_response {index} Your message</code>"
        )
        send_to_telegram(message)

# ✅ **Generate AI Response**
def generate_ai_response(prompt):
    """Calls OpenAI to generate a response with error handling."""
    url = "https://api.openai.com/v1/chat/completions"
    headers = {"Authorization": f"Bearer {OPENAI_API_KEY}", "Content-Type": "application/json"}
    data = {
        "model": "gpt-3.5-turbo",
        "messages": [{"role": "system", "content": "You are a customer support assistant for NextTradeWave.com, a CFD FX broker."}, {"role": "user", "content": prompt}]
    }
    response = requests.post(url, headers=headers, json=data)

    if response.status_code == 200:
        return html.escape(response.json()["choices"][0]["message"]["content"])
    else:
        return "⚠️ AI Response Unavailable."

# ✅ **Refine AI Response**
async def refine_response(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Improve an AI response based on user feedback."""
    args = context.args
    if not args or len(args) < 2:
        await update.message.reply_text("⚠️ Specify response number & improvement.\nExample: `/refine_response 2 Make it more polite.`")
        return

    response_index = args[0]
    user_feedback = " ".join(args[1:])

    if response_index not in ai_responses:
        await update.message.reply_text("⚠️ Invalid response number.")
        return

    full_prompt = f"Refine this response: {ai_responses[response_index]}\nUser Feedback: {user_feedback}"
    improved_response = generate_ai_response(full_prompt)
    ai_responses[response_index] = improved_response

    await update.message.reply_text(f"🤖 <b>Improved AI Response:</b>\n{improved_response}", parse_mode="HTML")

# ✅ **Help Command**
async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """List all available bot commands."""
    help_text = (
        "<b>📜 Available Commands:</b>\n"
        "/fetch_emails - Fetch latest unread emails\n"
        "/fetch_recent X - Fetch all emails in the last X hours\n"
        "/suggest_response <email_number> <your message> - Generate AI response\n"
        "/refine_response <response_number> <your improvement> - Improve AI response\n"
        "/help - Show this help menu"
    )
    await update.message.reply_text(help_text, parse_mode="HTML")

# ✅ **Auto-Fetch Emails Every 5 Minutes**
async def auto_fetch_emails():
    while True:
        await send_email_to_telegram()
        await asyncio.sleep(300)

if __name__ == "__main__":
    telegram_app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()
    telegram_app.add_handler(CommandHandler("fetch_emails", send_email_to_telegram))
    telegram_app.add_handler(CommandHandler("help", help_command))
    telegram_app.run_polling()
    asyncio.run(auto_fetch_emails())
    flask_app.run(host="0.0.0.0", port=8080)
