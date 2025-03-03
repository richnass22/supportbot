import os
import requests
import openai
from flask import Flask, jsonify
from msal import ConfidentialClientApplication
from telegram import Bot
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Debugging: Print environment variables
print("🔹 CLIENT_ID:", os.getenv("CLIENT_ID"))
print("🔹 TENANT_ID:", os.getenv("TENANT_ID"))

if not os.getenv("CLIENT_SECRET"):
    print("❌ ERROR: CLIENT_SECRET is missing!")
    exit()

# Environment Variables
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

# Initialize Flask App
app = Flask(__name__)

# Initialize OpenAI API
openai.api_key = OPENAI_API_KEY

# Initialize Telegram Bot
telegram_bot = Bot(token=TELEGRAM_BOT_TOKEN)

def get_access_token():
    """ Authenticate with Microsoft and get an access token """
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )

    token_response = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    
    if "access_token" in token_response:
        print("✅ Authentication successful!")
        return token_response["access_token"]
    else:
        print("❌ Authentication failed:", token_response)
        return None

def fetch_unread_emails():
    """ Fetch unread emails from Outlook """
    access_token = get_access_token()
    if not access_token:
        return []

    url = "https://graph.microsoft.com/v1.0/me/messages?$filter=isRead eq false"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        return response.json().get("value", [])
    else:
        print("❌ Failed to fetch emails:", response.json())
        return []

def generate_response(prompt):
    """ Use OpenAI to generate a response """
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[{"role": "user", "content": prompt}]
        )
        return response["choices"][0]["message"]["content"]
    except Exception as e:
        print("❌ OpenAI Error:", e)
        return "I'm sorry, I couldn't process your request."

def send_telegram_message(message):
    """ Send a message via Telegram bot """
    telegram_bot.send_message(chat_id=TELEGRAM_CHAT_ID, text=message)

@app.route('/process-emails', methods=['GET'])
def process_emails():
    """ Process unread emails and send replies """
    emails = fetch_unread_emails()
    if not emails:
        return jsonify({"message": "No new unread emails"}), 200

    for email in emails:
        subject = email.get("subject", "No Subject")
        body = email.get("bodyPreview", "No Content")

        response = generate_response(f"Subject: {subject}\n\n{body}")

        send_telegram_message(f"📩 **New Email**\n🔹 **Subject:** {subject}\n💬 **Response:** {response}")

    return jsonify({"message": "Processed emails"}), 200

@app.route('/')
def home():
    return "✅ Support Bot is Running!", 200

if __name__ == '__main__':
import os

port = int(os.environ.get("PORT", 5000))  # Get Railway's PORT or default to 5000
app.run(host="0.0.0.0", port=port)

