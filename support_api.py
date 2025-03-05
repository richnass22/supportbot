import os
import requests
import asyncio
import re
import html
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from msal import ConfidentialClientApplication
from flask import Flask, jsonify
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes
from telegram import InlineKeyboardMarkup, InlineKeyboardButton

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
ai_responses = {}  # Stores AI suggestions for refining responses

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
def fetch_unread_emails(access_token):
    """Fetch unread emails, filtering out sent emails."""
    headers = {"Authorization": f"Bearer {access_token}"}
    query = "isRead eq false"
    response = requests.get(
        f"{EMAILS_URL}?$filter={query}&$orderby=receivedDateTime desc",
        headers=headers
    )
    if response.status_code == 200:
        emails = response.json().get("value", [])
        return [email for email in emails if email.get("from", {}).get("emailAddress", {}).get("address") != EMAIL_ADDRESS]
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
async def send_email_to_telegram(context: ContextTypes.DEFAULT_TYPE):
    """Fetch unread emails, clean the content, and send them to Telegram."""
    print("📥 Fetching emails...")
    access_token = get_access_token()

    if access_token:
        emails = fetch_unread_emails(access_token)

        if emails:
            print(f"✅ {len(emails)} unread emails found.")
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

                # Format message for better readability
                message = (
                    f"<b>📩 New Email Received</b> [#{index}]\n"
                    f"📌 <b>From:</b> {sender_name} ({sender_email})\n"
                    f"📌 <b>Subject:</b> {subject}\n"
                    f"🕒 <b>Received:</b> {received_time}\n"
                    f"📝 <b>Preview:</b> {body_text[:500]}...\n\n"
                    f"✍️ Reply with: <code>/suggest_response {index} Your message</code>"
                )

                print(f"📤 Sending email {index} to Telegram: {subject}")  # Debug Log

                # Use context.bot.send_message properly
                await context.bot.send_message(
                    chat_id=TELEGRAM_CHAT_ID,
                    text=message,
                    parse_mode="HTML",
                    disable_web_page_preview=True
                )
        else:
            print("📭 No unread emails found.")  # Debug Log
            await context.bot.send_message(chat_id=TELEGRAM_CHAT_ID, text="<b>📭 No new unread emails found.</b>", parse_mode="HTML")
            
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
    return html.escape(response.json()["choices"][0]["message"]["content"]) if response.status_code == 200 else "⚠️ AI Response Unavailable."

# ✅ **Fix Telegram Commands**
async def fetch_emails_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Trigger email fetch via Telegram command."""
    print("📥 Received /fetch_emails command.")
    await context.bot.send_message(chat_id=update.effective_chat.id, text="📬 Fetching latest unread emails...")
    await send_email_to_telegram(context)  # Pass context explicitly

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

    # Store response to allow refinement
    ai_responses[email_index] = ai_response

    await update.message.reply_text(f"🤖 <b>AI Suggested Reply:</b>\n{ai_response}\n\n🔄 Reply with: `/improve_response {email_index} Your adjustment`", parse_mode="HTML")

async def improve_response(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Allow refining AI responses interactively."""
    args = context.args
    if not args or len(args) < 2:
        await update.message.reply_text("⚠️ Specify an email number & improvement.\nExample: `/improve_response 2 Make it more formal`")
        return

    email_index = args[0]  
    improvement = " ".join(args[1:])

    if email_index not in ai_responses:
        await update.message.reply_text("⚠️ No AI response found for this email. Use `/suggest_response` first.")
        return

    full_prompt = f"Previous AI response: {ai_responses[email_index]}\n\nUser Improvement: {improvement}\n\nGenerate a refined version."
    new_response = generate_ai_response(full_prompt)

    ai_responses[email_index] = new_response

    await update.message.reply_text(f"🤖 <b>Refined AI Response:</b>\n{new_response}", parse_mode="HTML")

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Send a help message listing all available commands."""
    help_text = (
        "<b>🛠 SupportBot Commands:</b>\n\n"
        "📩 <code>/fetch_emails</code> - Fetch the latest unread emails.\n"
        "📜 <code>/fetch_recent X</code> - Fetch emails from the last X hours.\n"
        "✍️ <code>/suggest_response EMAIL_ID Your message</code> - Generate an AI response for a specific email.\n"
        "🔄 <code>/improve EMAIL_ID Your message</code> - Refine an AI response with additional details.\n"
        "⏳ <code>/set_auto_fetch 5</code> - Automatically fetch emails every 5 minutes.\n"
        "📖 <code>/help</code> - Show this help message."
    )

    try:
        await update.message.reply_text(help_text, parse_mode="HTML")
    except Exception as e:
        print(f"❌ Error sending help message: {e}")

if __name__ == "__main__":
    telegram_app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()
    telegram_app.add_handler(CommandHandler("fetch_emails", fetch_emails_command))
    telegram_app.add_handler(CommandHandler("suggest_response", suggest_response))
    telegram_app.add_handler(CommandHandler("improve_response", improve_response))
    telegram_app.add_handler(CommandHandler("help", help_command))

    telegram_app.run_polling()
    flask_app.run(host="0.0.0.0", port=8080)
