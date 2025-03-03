import os
import requests
from flask import Flask, request, jsonify
from msal import ConfidentialClientApplication
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters

# Load environment variables
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

# Initialize Flask App
app = Flask(__name__)

# Function to authenticate with Microsoft
def get_access_token():
    try:
        app_auth = ConfidentialClientApplication(
            CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            client_credential=CLIENT_SECRET
        )
        token_response = app_auth.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" in token_response:
            return token_response["access_token"]
        else:
            return None
    except Exception as e:
        print(f"❌ Authentication failed: {str(e)}")
        return None

# Function to fetch unread emails
def fetch_unread_emails():
    access_token = get_access_token()
    if not access_token:
        return {"error": "Failed to authenticate"}

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json"
    }
    url = f"https://graph.microsoft.com/v1.0/users/{EMAIL_ADDRESS}/messages?$filter=isRead eq false"
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        emails = response.json().get("value", [])
        return emails
    else:
        return {"error": response.text}

# Route to process emails
@app.route("/process-emails", methods=["GET"])
def process_emails():
    emails = fetch_unread_emails()
    return jsonify(emails)

# Telegram Bot Setup
telegram_app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()

async def start(update: Update, context):
    await update.message.reply_text("Hello! I'm your support bot. How can I assist you?")

async def handle_message(update: Update, context):
    user_message = update.message.text
    await update.message.reply_text(f"You said: {user_message}")

telegram_app.add_handler(CommandHandler("start", start))
telegram_app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

# Run Flask & Telegram bot
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    telegram_app.run_polling()  # Starts the Telegram bot
    app.run(host="0.0.0.0", port=port)
