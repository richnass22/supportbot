import os
import requests
import asyncio
import html
import threading
from msal import ConfidentialClientApplication
from flask import Flask, jsonify
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes

# 🔹 Load environment variables
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")  # Ensure this is set correctly!
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

# 🔹 Fetch Emails
def fetch_emails(access_token):
    """Fetch emails from Microsoft Graph API"""
    headers = {"Authorization": f"Bearer {access_token}"}
    
    response = requests.get(EMAILS_URL, headers=headers)

    if response.status_code == 200:
        print("📩 Emails fetched successfully.")
        return response.json().get("value", [])
    else:
        print(f"❌ Error fetching emails: {response.json()}")
        return None

# 🔹 Send Message to Telegram
def send_to_telegram(message):
    """Send a message to Telegram with escaped HTML characters"""
    telegram_url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    
    # Escape special characters for Telegram
    escaped_message = html.escape(message)

    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": escaped_message,
        "parse_mode": "HTML"
    }

    response = requests.post(telegram_url, json=payload)

    if response.status_code == 200:
        print("📤 Message sent to Telegram successfully.")
    else:
        print(f"❌ Error sending to Telegram: {response.json()}")

# 🔹 Async Function for Email Processing
async def send_email_to_telegram():
    """Fetch emails and send them to Telegram"""
    access_token = get_access_token()
    
    if access_token:
        emails = fetch_emails(access_token)
        
        if emails:
            email_store.clear()  # Reset previous emails
            for index, email in enumerate(emails[:5], start=1):  # Process top 5 emails
                subject = email.get("subject", "No Subject")
                sender_name = email.get("from", {}).get("emailAddress", {}).get("name", "Unknown Sender")
                sender_email = email.get("from", {}).get("emailAddress", {}).get("address", "Unknown Email")
                body_preview = email.get("bodyPreview", "No Preview Available")

                # Store email data in dictionary for future reference
                email_store[str(index)] = {"sender": sender_name, "subject": subject, "body": body_preview}

                # Escape special characters for Telegram
                subject = html.escape(subject)
                sender_name = html.escape(sender_name)
                sender_email = html.escape(sender_email)
                body_preview = html.escape(body_preview)

                message = (
                    f"📩 <b>New Email Received</b> [#{index}]\n"
                    f"📌 <b>From:</b> {sender_name} ({sender_email})\n"
                    f"📌 <b>Subject:</b> {subject}\n"
                    f"📌 <b>Preview:</b> {body_preview}\n"
                    f"✍️ Reply with: <code>/suggest_response {index} Your message</code>"
                )

                send_to_telegram(message)
        else:
            print("⚠️ No new emails found.")

# 🔹 Flask Route to Trigger Email Fetching
@flask_app.route("/process-emails", methods=["GET"])
def process_emails():
    """Trigger the email fetch function."""
    asyncio.run(send_email_to_telegram())  # ✅ Fixed: Ensures an event loop is running
    return jsonify({"message": "Fetching emails... Check your Telegram!"})

# 🔹 Generate AI Response
def generate_ai_response(prompt):
    """Calls OpenAI to generate a response."""
    url = "https://api.openai.com/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "Content-Type": "application/json"
    }
    data = {
        "model": "gpt-3.5-turbo",
        "messages": [{"role": "system", "content": "You are a professional customer support assistant."},
                     {"role": "user", "content": prompt}]
    }
    response = requests.post(url, headers=headers, json=data)
    
    if response.status_code == 200:
        return response.json()["choices"][0]["message"]["content"]
    else:
        return f"⚠️ Error generating AI response: {response.text}"

# === 🤖 TELEGRAM BOT COMMANDS === #
async def fetch_emails_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Trigger email fetch via Telegram command."""
    await context.bot.send_message(chat_id=update.effective_chat.id, text="📬 Fetching emails...")
    await send_email_to_telegram()

async def suggest_response(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Generate AI response based on selected email."""
    args = context.args
    if not args or len(args) < 2:
        await update.message.reply_text("⚠️ Please specify an email number and message.\nExample: `/suggest_response 2 Please be polite`")
        return

    email_index = args[0]  # First argument should be the email number
    user_message = " ".join(args[1:])  # The rest is the message

    if email_index not in email_store:
        await update.message.reply_text("⚠️ Invalid email number. Use `/fetch_emails` to get valid email IDs.")
        return

    email_data = email_store[email_index]
    full_prompt = f"Email Subject: {email_data['subject']}\n\nEmail Body: {email_data['body']}\n\nUser Instruction: {user_message}"

    ai_response = generate_ai_response(full_prompt)
    await update.message.reply_text(f"🤖 AI Suggested Reply:\n{ai_response}")

# Run Flask Server & Telegram Bot Together
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    telegram_thread = threading.Thread(target=lambda: ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build().run_polling(), daemon=True)
    telegram_thread.start()
    flask_app.run(host="0.0.0.0", port=port)
