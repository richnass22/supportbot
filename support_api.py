import os
import asyncio
import requests
from msal import ConfidentialClientApplication
from flask import Flask, jsonify
from telegram.ext import ApplicationBuilder, CommandHandler
import openai  # Ensure you have the OpenAI API installed

# Load environment variables
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

# Validate critical environment variables
missing_vars = []
for var_name, var_value in {
    "CLIENT_ID": CLIENT_ID,
    "CLIENT_SECRET": CLIENT_SECRET,
    "TENANT_ID": TENANT_ID,
    "EMAIL_ADDRESS": EMAIL_ADDRESS,
    "OPENAI_API_KEY": OPENAI_API_KEY,
    "TELEGRAM_BOT_TOKEN": TELEGRAM_BOT_TOKEN,
    "TELEGRAM_CHAT_ID": TELEGRAM_CHAT_ID,
}.items():
    if not var_value:
        missing_vars.append(var_name)

if missing_vars:
    raise ValueError(f"❌ Missing required environment variables: {', '.join(missing_vars)}")

# Print confirmation of loaded variables
print(f"✅ CLIENT_ID Loaded: {CLIENT_ID[:5]}...")
print(f"✅ TENANT_ID Loaded: {TENANT_ID[:5]}...")
print(f"✅ TELEGRAM_BOT_TOKEN Loaded: {TELEGRAM_BOT_TOKEN[:5]}...")

# Authenticate with Microsoft Graph API
auth_app = ConfidentialClientApplication(
    CLIENT_ID, 
    authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    client_credential=CLIENT_SECRET
)

token_response = auth_app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
if "access_token" not in token_response:
    raise ValueError(f"❌ Authentication failed: {token_response}")

print("✅ Authentication successful!")

# Initialize Flask App
flask_app = Flask(__name__)

@flask_app.route("/")
def home():
    return jsonify({"message": "Support Bot is running!"})

# Initialize Telegram Bot
telegram_app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()

async def send_email_to_telegram():
    """Fetches emails and sends them to Telegram with AI-generated suggestions."""
    access_token = token_response["access_token"]
    url = f"https://graph.microsoft.com/v1.0/me/messages?$top=1"
    headers = {"Authorization": f"Bearer {access_token}"}

    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        emails = response.json().get("value", [])
        if not emails:
            message = "📭 No new emails found."
        else:
            email = emails[0]
            sender = email['from']['emailAddress']['address']
            subject = email.get("subject", "No Subject")
            body_preview = email.get("bodyPreview", "No Content")

            # Generate AI response suggestion
            openai.api_key = OPENAI_API_KEY
            ai_response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "system", "content": "Suggest a professional response based on the email details."},
                          {"role": "user", "content": f"Sender: {sender}\nSubject: {subject}\nBody: {body_preview}"}]
            )

            suggested_reply = ai_response["choices"][0]["message"]["content"]

            message = (
                f"📩 *New Email Received!*\n"
                f"👤 *From:* {sender}\n"
                f"📌 *Subject:* {subject}\n"
                f"📝 *Preview:* {body_preview}\n\n"
                f"💡 *Suggested Reply:* {suggested_reply}"
            )
    else:
        message = "⚠️ Error fetching emails."

    # Send message to Telegram
    telegram_url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    data = {"chat_id": TELEGRAM_CHAT_ID, "text": message, "parse_mode": "Markdown"}
    requests.post(telegram_url, data=data)

@flask_app.route("/process-emails", methods=["GET"])
def process_emails():
    """Triggers email processing and sends to Telegram."""
    asyncio.create_task(send_email_to_telegram())  # ✅ Fixed from `.loop.create_task()`
    return jsonify({"message": "✅ Email processing started! Check Telegram."})

# Telegram Command: Fetch Emails
async def fetch_emails(update, context):
    """Allows users to trigger email fetching via Telegram command."""
    await send_email_to_telegram()
    await update.message.reply_text("✅ Fetching emails... Check your messages!")

telegram_app.add_handler(CommandHandler("fetch_emails", fetch_emails))

# Run Flask & Telegram Bot
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))  # Default to 8080
    telegram_app.run_polling()  # Ensures the Telegram bot runs
    flask_app.run(host="0.0.0.0", port=port)
