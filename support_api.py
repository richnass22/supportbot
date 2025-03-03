import os
import requests
import asyncio
from msal import ConfidentialClientApplication
from flask import Flask, jsonify
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes

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

# Print confirmation of loaded variables (partial for security)
print(f"✅ CLIENT_ID Loaded: {CLIENT_ID[:5]}...")
print(f"✅ TENANT_ID Loaded: {TENANT_ID[:5]}...")
print(f"✅ TELEGRAM_BOT_TOKEN Loaded: {TELEGRAM_BOT_TOKEN[:5]}...")

# Authenticate with Microsoft Graph API
app = ConfidentialClientApplication(
    CLIENT_ID, 
    authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    client_credential=CLIENT_SECRET
)

token_response = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
if "access_token" not in token_response:
    raise ValueError(f"❌ Authentication failed: {token_response}")

print("✅ Authentication successful!")

# Flask App Setup
flask_app = Flask(__name__)

@flask_app.route("/")
def home():
    return jsonify({"message": "Support Bot is running!"})

# === 📩 FETCH EMAILS & SEND TO TELEGRAM === #
async def send_email_to_telegram():
    """Fetches emails and sends them to Telegram with AI-generated suggestions."""
    access_token = token_response["access_token"]
    url = f"https://graph.microsoft.com/v1.0/me/messages?$top=1"
    headers = {"Authorization": f"Bearer {access_token}"}

    response = requests.get(url, headers=headers)
    
    # Log the response for debugging
    print("📥 Email Fetch Response:", response.status_code, response.text)

    if response.status_code == 200:
        emails = response.json().get("value", [])
        if not emails:
            message = "📭 No new emails found."
        else:
            email = emails[0]
            sender = email['from']['emailAddress']['address']
            subject = email.get("subject", "No Subject")
            body_preview = email.get("bodyPreview", "No Content")

            message = (
                f"📩 *New Email Received!*\n"
                f"👤 *From:* {sender}\n"
                f"📌 *Subject:* {subject}\n"
                f"📝 *Preview:* {body_preview}\n\n"
                f"✍️ Reply with: /suggest_response"
            )
    else:
        message = f"⚠️ Error fetching emails: {response.text}"

    # Send message to Telegram
    telegram_url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    data = {"chat_id": TELEGRAM_CHAT_ID, "text": message, "parse_mode": "Markdown"}
    requests.post(telegram_url, data=data)

@flask_app.route("/process-emails", methods=["GET"])
def process_emails():
    """Trigger the email fetch function."""
    asyncio.create_task(send_email_to_telegram())
    return jsonify({"message": "Fetching emails... Check your Telegram!"})

# === 🤖 TELEGRAM BOT COMMANDS === #
async def fetch_emails(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Trigger email fetch via Telegram command."""
    await context.bot.send_message(chat_id=update.effective_chat.id, text="📬 Fetching emails...")
    await send_email_to_telegram()

async def suggest_response(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Ask AI to generate a response based on user input."""
    user_input = " ".join(context.args)
    
    if not user_input:
        await context.bot.send_message(
            chat_id=update.effective_chat.id, 
            text="✍️ Please provide some pointers for the AI.\nExample: /suggest_response The customer is asking about refund policies."
        )
        return

    # Send to OpenAI GPT API for response generation
    ai_response = generate_ai_response(user_input)
    
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text=f"🤖 AI Suggested Reply:\n{ai_response}"
    )

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

# === 🤖 TELEGRAM BOT SETUP === #
try:
    telegram_app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()
    
    telegram_app.add_handler(CommandHandler("fetch_emails", fetch_emails))
    telegram_app.add_handler(CommandHandler("suggest_response", suggest_response))
    
    print("✅ Telegram bot initialized successfully!")
except Exception as e:
    raise ValueError(f"❌ Telegram bot initialization failed: {e}")

# Run Flask Server
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))  # Default to 8080
    flask_app.run(host="0.0.0.0", port=port)
