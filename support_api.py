import os
import requests
import asyncio
import re
import html
import logging
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from msal import ConfidentialClientApplication
from flask import Flask, jsonify
from telegram import Update, InlineKeyboardMarkup, InlineKeyboardButton
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes
from telegram.error import BadRequest

# 🔹 Enable Logging
logging.basicConfig(level=logging.INFO)

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
ai_responses = {}

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
    return response.json().get("access_token") if response.status_code == 200 else None

# 🔹 Fetch Unread Emails
def fetch_unread_emails(access_token, hours=24):
    """Fetch unread emails, filtering out sent emails, and retrieving last X hours."""
    headers = {"Authorization": f"Bearer {access_token}"}
    since_time = (datetime.utcnow() - timedelta(hours=int(hours))).isoformat() + "Z"
    query = f"isRead eq false and receivedDateTime ge {since_time}"

    response = requests.get(
        f"{EMAILS_URL}?$filter={query}&$orderby=receivedDateTime desc",
        headers=headers
    )
    if response.status_code == 200:
        emails = response.json().get("value", [])
        return [email for email in emails if email.get("from", {}).get("emailAddress", {}).get("address") != EMAIL_ADDRESS]
    return None

# 🔹 Send Messages to Telegram (Using HTML Mode)
async def send_to_telegram(context: ContextTypes.DEFAULT_TYPE, message):
    """Send a well-formatted message to Telegram using HTML mode."""
    try:
        await context.bot.send_message(
            chat_id=TELEGRAM_CHAT_ID,
            text=html.escape(message),
            parse_mode="HTML",
            disable_web_page_preview=True
        )
        logging.info("📤 Message sent to Telegram successfully.")
    except BadRequest as e:
        logging.error(f"❌ Error sending to Telegram: {e}")

# 🔹 Async Function for Email Processing
async def send_email_to_telegram(context: ContextTypes.DEFAULT_TYPE, hours=24):
    """Fetch unread emails, clean the content, and send them to Telegram."""
    logging.info("📥 Fetching emails...")
    access_token = get_access_token()

    if access_token:
        emails = fetch_unread_emails(access_token, hours)

        if emails:
            logging.info(f"✅ {len(emails)} unread emails found.")
            email_store.clear()

            for index, email in enumerate(emails[:10], start=1):  # Fetch up to 10 emails
                subject = email.get("subject", "No Subject")
                sender_name = email.get("from", {}).get("emailAddress", {}).get("name", "Unknown Sender")
                sender_email = email.get("from", {}).get("emailAddress", {}).get("address", "Unknown Email")
                received_time = email.get("receivedDateTime", "Unknown Time")
                body_html = email.get("body", {}).get("content", "No Preview Available")

                soup = BeautifulSoup(body_html, "html.parser")
                body_text = soup.get_text()

                email_store[str(index)] = {
                    "sender": sender_name,
                    "subject": subject,
                    "body": body_text
                }

                message = (
                    f"<b>📩 New Email Received</b> [#{index}]\n"
                    f"📌 <b>From:</b> {sender_name} ({sender_email})\n"
                    f"📌 <b>Subject:</b> {subject}\n"
                    f"🕒 <b>Received:</b> {received_time}\n"
                    f"📝 <b>Preview:</b> {body_text[:500]}...\n\n"
                    f"✍️ Reply with: <code>/suggest_response {index} Your message</code>"
                )

                await send_to_telegram(context, message)
        else:
            await send_to_telegram(context, "<b>📭 No new unread emails found.</b>")

# ✅ **Fix Telegram Commands**
async def fetch_emails_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Trigger email fetch via Telegram command."""
    await update.message.reply_text("📬 Fetching latest unread emails...")
    await send_email_to_telegram(context)

async def fetch_recent_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Fetch emails from the last X hours."""
    args = context.args
    if not args or not args[0].isdigit():
        await update.message.reply_text("⚠️ Use `/fetch_recent X` where X is the number of hours.")
        return

    hours = int(args[0])
    await update.message.reply_text(f"📬 Fetching all unread emails from the last {hours} hours...")

    access_token = get_access_token()
    if not access_token:
        await update.message.reply_text("❌ Error: Could not retrieve access token.")
        return

    emails = fetch_unread_emails(access_token, hours)
    if not emails:
        await update.message.reply_text("📭 No unread emails found.")
        return

    for index, email in enumerate(emails, start=1):  # Now processes ALL emails
        subject = email.get("subject", "No Subject")
        sender_name = email.get("from", {}).get("emailAddress", {}).get("name", "Unknown Sender")
        sender_email = email.get("from", {}).get("emailAddress", {}).get("address", "Unknown Email")
        received_time = email.get("receivedDateTime", "Unknown Time")
        body_html = email.get("body", {}).get("content", "No Preview Available")

        soup = BeautifulSoup(body_html, "html.parser")
        body_text = soup.get_text()

        email_store[str(index)] = {
            "sender": sender_name,
            "subject": subject,
            "body": body_text
        }

        message = (
            f"<b>📩 New Email Received</b> [#{index}]\n"
            f"📌 <b>From:</b> {sender_name} ({sender_email})\n"
            f"📌 <b>Subject:</b> {subject}\n"
            f"🕒 <b>Received:</b> {received_time}\n"
            f"📝 <b>Preview:</b> {body_text[:500]}...\n\n"
            f"✍️ Reply with: <code>/suggest_response {index} Your message</code>"
        )

        try:
            await update.message.reply_text(message, parse_mode="HTML", disable_web_page_preview=True)
        except Exception as e:
            print(f"❌ Error sending message to Telegram: {e}")

    await update.message.reply_text(f"✅ Displayed all unread emails from the last {hours} hours.")

    hours = int(args[0])
    await update.message.reply_text(f"📬 Fetching emails from the last {hours} hours...")
    await send_email_to_telegram(context, hours)

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Send a help message listing all available commands."""
    help_text = (
        "<b>🛠 SupportBot Commands:</b>\n\n"
        "📩 <code>/fetch_emails</code> - Fetch the latest unread emails.\n"
        "📜 <code>/fetch_recent X</code> - Fetch emails from the last X hours.\n"
        "✍️ <code>/suggest_response EMAIL_ID Your message</code> - Generate an AI response.\n"
        "🔄 <code>/improve_response EMAIL_ID Your message</code> - Improve an AI response.\n"
        "⏳ <code>/set_auto_fetch 5</code> - Automatically fetch emails every 5 minutes.\n"
        "📖 <code>/help</code> - Show this help message."
    )

    await update.message.reply_text(help_text, parse_mode="HTML")

if __name__ == "__main__":
    telegram_app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()
    telegram_app.add_handler(CommandHandler("fetch_emails", fetch_emails_command))
    telegram_app.add_handler(CommandHandler("fetch_recent", fetch_recent_command))
    telegram_app.add_handler(CommandHandler("help", help_command))

    telegram_app.run_polling()
    flask_app.run(host="0.0.0.0", port=8080)
