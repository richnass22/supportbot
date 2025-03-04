import os
import requests
import asyncio
import html
import threading
from datetime import datetime, timedelta
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

# 🔹 Send Message to Telegram (Better Formatting)
def send_to_telegram(message):
    """Send a well-formatted message to Telegram."""
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
    """Fetch unread emails and send them to Telegram."""
    access_token = get_access_token()
    
    if access_token:
        emails = fetch_unread_emails(access_token, hours)
        
        if emails:
            email_store.clear()  # Reset previous emails
            for index, email in enumerate(emails[:5], start=1):  # Process top 5 emails
                subject = email.get("subject", "No Subject")
                sender_name = email.get("from", {}).get("emailAddress", {}).get("name", "Unknown Sender")
                sender_email = email.get("from", {}).get("emailAddress", {}).get("address", "Unknown Email")
                body_preview = email.get("bodyPreview", "No Preview Available")
                received_time = email.get("receivedDateTime", "Unknown Time")

                # Store email data in dictionary for future reference
                email_store[str(index)] = {
                    "sender": sender_name,
                    "subject": subject,
                    "body": body_preview
                }

                # Format message for better readability
                message = (
                    f"📩 <b>New Email Received</b> [#{index}]\n"
                    f"📌 <b>From:</b> {sender_name} ({sender_email})\n"
                    f"📌 <b>Subject:</b> {subject}\n"
                    f"🕒 <b>Received:</b> {received_time}\n"
                    f"📝 <b>Preview:</b> {body_preview[:500]}...\n\n"
                    f"✍️ Reply with: <code>/suggest_response {index} Your message</code>"
                )

                send_to_telegram(message)
        else:
            send_to_telegram("📭 No new unread emails found.")

# 🔹 Flask Route to Trigger Email Fetching
@flask_app.route("/process-emails", methods=["GET"])
def process_emails():
    """Trigger the email fetch function for last 24 hours."""
    asyncio.run(send_email_to_telegram(24))  
    return jsonify({"message": "Fetching last 24 hours of emails... Check Telegram!"})

# === 🤖 TELEGRAM BOT COMMANDS === #
async def fetch_emails_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Trigger email fetch via Telegram command."""
    await context.bot.send_message(chat_id=update.effective_chat.id, text="📬 Fetching latest unread emails...")
    await send_email_to_telegram()

async def fetch_recent_emails(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Fetch unread emails from the last X hours."""
    args = context.args
    hours = args[0] if args else "24"  # Default to last 24 hours if no argument given

    await context.bot.send_message(chat_id=update.effective_chat.id, text=f"📬 Fetching unread emails from the last {hours} hours...")
    await send_email_to_telegram(hours)

# ✅ **Start Telegram Bot Properly**
def start_telegram_bot():
    """Runs the Telegram bot properly in the main thread"""
    telegram_app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()
    telegram_app.add_handler(CommandHandler("fetch_emails", fetch_emails_command))
    telegram_app.add_handler(CommandHandler("fetch_recent", fetch_recent_emails))

    print("✅ Telegram bot initialized successfully!")
    telegram_app.run_polling()

# 🔹 Run Flask Server & Telegram Bot Together
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    start_telegram_bot()  # Start Telegram bot
    flask_app.run(host="0.0.0.0", port=port)  # Start Flask server
