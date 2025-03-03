import os
import requests
import time
from msal import ConfidentialClientApplication
from flask import Flask, jsonify
from telegram import Bot
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes

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

print(f"✅ CLIENT_ID Loaded: {CLIENT_ID[:5]}...")
print(f"✅ TENANT_ID Loaded: {TENANT_ID[:5]}...")
print(f"✅ TELEGRAM_BOT_TOKEN Loaded: {TELEGRAM_BOT_TOKEN[:5]}...")

# Authenticate with Microsoft Graph API
app = ConfidentialClientApplication(
    CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    client_credential=CLIENT_SECRET,
)

token_response = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
if "access_token" not in token_response:
    raise ValueError(f"❌ Authentication failed: {token_response}")

print("✅ Authentication successful!")

# Flask App Setup
flask_app = Flask(__name__)

# Telegram Bot Setup
bot = Bot(token=TELEGRAM_BOT_TOKEN)
telegram_app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()

# Dictionary to store user input
user_feedback = {}

### 📨 Fetch Unread Emails
def fetch_unread_emails():
    access_token = token_response["access_token"]
    headers = {"Authorization": f"Bearer {access_token}"}
    url = "https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages?$filter=isRead eq false&$top=1"

    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        emails = response.json().get("value", [])
        if emails:
            email = emails[0]
            sender = email["from"]["emailAddress"]["address"]
            subject = email["subject"]
            body = email["body"]["content"]
            return {"sender": sender, "subject": subject, "body": body}
        else:
            return None
    else:
        print("❌ Error fetching emails:", response.json())
        return None

### 🤖 Generate AI Response
def generate_ai_response(email_body, user_feedback=""):
    prompt = f"""
    You are an AI assistant handling emails. Here is an unread email:

    --- EMAIL CONTENT ---
    {email_body}

    --- USER INPUT ---
    {user_feedback}

    Based on the email content and user input, draft a professional response.
    """

    headers = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": "gpt-4",
        "messages": [{"role": "system", "content": prompt}]
    }

    response = requests.post("https://api.openai.com/v1/chat/completions", json=payload, headers=headers)
    return response.json().get("choices", [{}])[0].get("message", {}).get("content", "⚠️ AI Response Not Generated.")

### 🚀 Send Email to Telegram
async def send_email_to_telegram():
    email = fetch_unread_emails()
    if not email:
        await bot.send_message(chat_id=TELEGRAM_CHAT_ID, text="📭 No unread emails found.")
        return

    ai_response = generate_ai_response(email["body"])

    message = (
        f"📩 *New Email Received!*\n\n"
        f"👤 *From:* {email['sender']}\n"
        f"📌 *Subject:* {email['subject']}\n\n"
        f"📨 *Email Content:*\n{email['body'][:500]}...\n\n"
        f"🤖 *AI Suggested Reply:*\n{ai_response}\n\n"
        f"✏️ *Reply with feedback to refine the response:* `/feedback [your thoughts]`"
    )

    await bot.send_message(chat_id=TELEGRAM_CHAT_ID, text=message, parse_mode="Markdown")
    user_feedback[TELEGRAM_CHAT_ID] = email["body"]  # Store email content for feedback processing

### ✏️ Handle Telegram Feedback
async def handle_feedback(update, context: ContextTypes.DEFAULT_TYPE):
    feedback_text = " ".join(context.args)
    if not feedback_text:
        await update.message.reply_text("⚠️ Please provide feedback. Example:\n`/feedback Make it more polite`")
        return

    email_body = user_feedback.get(update.message.chat_id, "")
    if not email_body:
        await update.message.reply_text("❌ No email found to refine response.")
        return

    # Generate refined AI response
    refined_response = generate_ai_response(email_body, feedback_text)
    await update.message.reply_text(f"✅ *Updated AI Response:*\n{refined_response}", parse_mode="Markdown")

# Add Telegram Command Handlers
telegram_app.add_handler(CommandHandler("fetch_emails", lambda update, ctx: send_email_to_telegram()))
telegram_app.add_handler(CommandHandler("feedback", handle_feedback))

### 📬 Email Processing Route
@flask_app.route("/process-emails", methods=["GET"])
def process_emails():
    telegram_app.loop.create_task(send_email_to_telegram())
    return jsonify({"message": "✅ Email fetched & sent to Telegram!"})

# Run Flask Server
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    flask_app.run(host="0.0.0.0", port=port)
    telegram_app.run_polling()
