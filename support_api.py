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
        emails = response.json().get("value", [])

        # Debugging: Print the number of emails fetched
        print(f"📩 Debug: {len(emails)} unread emails fetched.")

        # Print first email for debugging
        if emails:
            print(f"📝 First email: {emails[0]}")

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
        "parse_mode": "MarkdownV2",
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
                    f"📩 *New Email Received* \\[#{index}\\]\n"
                    f"📌 *From:* {sender_name} \\({sender_email}\\)\n"
                    f"📌 *Subject:* {subject}\n"
                    f"🕒 *Received:* {received_time}\n"
                    f"📝 *Preview:* {body_preview[:500]}...\n\n"
                    f"✍️ Reply with: `/suggest_response {index} Your message`"
                )

                send_to_telegram(message)
        else:
            send_to_telegram("📭 *No new unread emails found.*")

# 🔹 Generate AI Response (With Fixes & Debugging)
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
            {"role": "system", "content": "You are a customer support bot for NextTradeWave.com, a CFD FX broker. Your responses should reflect this and avoid assuming another company’s identity."},
            {"role": "user", "content": prompt}
        ]
    }
    response = requests.post(url, headers=headers, json=data)

    if response.status_code == 200:
        return response.json()["choices"][0]["message"]["content"]
    else:
        return f"⚠️ AI Response Unavailable: {response.json().get('error', {}).get('message', 'Unknown error occurred.')}\nPlease check OpenAI API status or billing."

# === 🤖 TELEGRAM BOT COMMANDS === #
async def fetch_emails_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Trigger email fetch via Telegram command."""
    print("📥 Received /fetch_emails command.")
    await context.bot.send_message(chat_id=update.effective_chat.id, text="📬 Fetching latest unread emails...")
    await send_email_to_telegram()

async def suggest_response(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Generate AI response based on selected email."""
    args = context.args
    if not args or len(args) < 2:
        await update.message.reply_text("⚠️ Specify an email number & message.\nExample: `/suggest_response 2 Apologize for the delay`")
        return

    email_index = args[0]  # First argument should be the email number
    user_message = " ".join(args[1:])  # The rest is the message

    if email_index not in email_store:
        await update.message.reply_text("⚠️ Invalid email number. Use `/fetch_emails` first.")
        return

    email_data = email_store[email_index]
    full_prompt = f"Company: NextTradeWave.com (CFD FX Broker)\n\nEmail Subject: {email_data['subject']}\n\nEmail Body: {email_data['body']}\n\nUser Instruction: {user_message}"

    ai_response = generate_ai_response(full_prompt)

    await update.message.reply_text(f"🤖 *AI Suggested Reply:*\n{ai_response}", parse_mode="MarkdownV2")

def start_telegram_bot():
    telegram_app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()
    telegram_app.add_handler(CommandHandler("fetch_emails", fetch_emails_command))
    telegram_app.add_handler(CommandHandler("suggest_response", suggest_response))
    telegram_app.run_polling()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    start_telegram_bot()
    flask_app.run(host="0.0.0.0", port=port)
