import os
import requests
from msal import ConfidentialClientApplication
from flask import Flask, jsonify
from telegram.ext import ApplicationBuilder

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

# Process Emails Route
@flask_app.route("/process-emails", methods=["GET"])
def process_emails():
    return jsonify({"message": "Email processing will be implemented soon!"})

# Initialize Telegram Bot
try:
    telegram_app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()
    print("✅ Telegram bot initialized successfully!")
except Exception as e:
    raise ValueError(f"❌ Telegram bot initialization failed: {e}")

# Run Flask Server
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))  # Default to 8080
    flask_app.run(host="0.0.0.0", port=port)
