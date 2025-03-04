import os
import requests
import asyncio
import html
import threading
import re
from datetime import datetime, timedelta
from bs4 import BeautifulSoup  # For stripping HTML from emails
from msal import ConfidentialClientApplication
from flask import Flask, jsonify
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes

# 🔹 Load environment variables
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")  # Your company's email
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

# 🔹 Microsoft Graph API Endpoints
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
EMAILS_URL = f"https://graph.microsoft.com/v1.0/users/{EMAIL_ADDRESS}/messages"

# 🔹 Temporary Storage for Emails
email_store = {}

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
        print("✅ Access token retrieved successfully.")
        return response.json().get("access_token")
    else:
        print(f"❌ Error fetching token: {response.json()}")
        return None

# 🔹 Fetch Unread Emails (Filters out sent emails)
def fetch_unread_emails(access_token, hours=None):
    """Fetch unread emails from Microsoft Graph API, filtering out sent emails."""
    headers = {"Authorization": f"Bearer {access_token}"}
    
    # Base query to fetch unread emails
    query = "isRead eq false"
    
    # Filter by time range if specified
    if hours:
        time_filter = (datetime.utcnow() - timedelta(hours=int(hours))).isoformat() + "Z"
        query += f" and receivedDateTime ge {time_filter}"

    # API request with filtering
    response = requests.get(
        f"{EMAILS_URL}?$filter={query}&$orderby=receivedDateTime desc",
        headers=headers
    )

    if response.status_code == 200:
        print("📩 Unread emails fetched successfully.")
        emails = response.json().get("value", [])

        # Filter out outgoing emails sent by our company
        filtered_emails = [
            email for email in emails
            if email.get("from", {}).get("emailAddress", {}).get("address") != EMAIL_ADDRESS
        ]

        return filtered_emails
    else:
        print(f"❌ Error fetching unread emails: {response.json()}")
        return None

# 🔹 Send Message to Telegram (Uses HTML Mode for Stability)
def send_to_telegram(message):
    """Send a well-formatted message to Telegram using HTML mode."""
    telegram_url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"

    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": message,
        "parse_mode": "HTML",
        "disable_web_page_preview": True
    }

    response = requests.post(telegram_url, json=payload)

    if response.status_code == 200:
        print("📤 Message sent to Telegram successfully.")
    else:
        print(f"❌ Error sending to Telegram: {response.json()}")

# 🔹 Async Function for Email Processing
async def send_email_to_telegram(hours=None):
    """Fetch unread emails, clean the content, and send them to Telegram."""
    access_token = get_access_token()
    
    if access_token:
        emails = fetch_unread_emails(access_token, hours)
        
        if emails:
            email_store.clear()  # Reset previous emails
            for index, email in enumerate(emails[:5], start=1):  # Process top 5 emails
                subject = email.get("subject", "No Subject")
                sender_name = email.get("from", {}).get("emailAddress", {}).get("name", "Unknown Sender")
                sender_email = email.get("from", {}).get("emailAddress", {}).get("address", "Unknown Email")
                received_time = email.get("receivedDateTime", "Unknown Time")
                body_html = email.get("body", {}).get("content", "No Preview Available")

                # Convert HTML to plain text
                soup = BeautifulSoup(body_html, "html.parser")
                body_text = soup.get_text()

                # Store email data
                email_store[str(index)] = {
                    "sender": sender_name,
                    "subject": subject,
                    "body": body_text
                }

                # Limit message size for Telegram
                body_preview = body_text[:500] + "..." if len(body_text) > 500 else body_text

                # Format message for better readability
                message = (
                    f"<b>📩 New Email Received</b> [#{index}]\n"
                    f"📌 <b>From:</b> {sender_name} ({sender_email})\n"
                    f"📌 <b>Subject:</b> {subject}\n"
                    f"🕒 <b>Received:</b> {received_time}\n"
                    f"📝 <b>Preview:</b> {body_preview}\n\n"
                    f"✍️ Reply with: <code>/suggest_response {index} Your message</code>"
                )

                send_to_telegram(message)
        else:
            send_to_telegram("<b>📭 No new unread emails found.</b>")

# ✅ **Define AI Response Function**
def generate_ai_response(prompt):
    """Calls OpenAI to generate a response with error handling."""
    url = "https://api.openai.com/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "Content-Type": "application/json"
    }
    data = {
        "model": "gpt-3.5-turbo",
        "messages": [
            {"role": "system", "content": "You are a customer support assistant for NextTradeWave.com, a CFD FX broker."},
            {"role": "user", "content": prompt}
        ]
    }
    response = requests.post(url, headers=headers, json=data)

    if response.status_code == 200:
        return response.json()["choices"][0]["message"]["content"]
    else:
        error_msg = response.json().get("error", {}).get("message", "Unknown error")
        return f"⚠️ AI Response Unavailable: {error_msg}\nPlease check OpenAI API status or billing."

# ✅ **Define `/fetch_emails` Command**
async def fetch_emails_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Trigger email fetch via Telegram command."""
    print("📥 Received /fetch_emails command.")
    await context.bot.send_message(chat_id=update.effective_chat.id, text="📬 Fetching latest unread emails...")
    await send_email_to_telegram()

# ✅ **Define `/suggest_response` Command**
async def suggest_response(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Generate AI response based on selected email."""
    args = context.args
    if not args or len(args) < 2:
        await update.message.reply_text("⚠️ Specify an email number & message.\nExample: `/suggest_response 2 Apologize for the delay`")
        return

    email_index = args[0]  
    user_message = " ".join(args[1:])

    email_data = email_store.get(email_index)
    if not email_data:
        await update.message.reply_text("⚠️ Invalid email number. Use `/fetch_emails` first.")
        return

    full_prompt = f"Company: NextTradeWave.com\n\nEmail Subject: {email_data['subject']}\n\nEmail Body: {email_data['body']}\n\nUser Instruction: {user_message}"

    ai_response = generate_ai_response(full_prompt)

    await update.message.reply_text(f"🤖 <b>AI Suggested Reply:</b>\n{ai_response}", parse_mode="HTML")
