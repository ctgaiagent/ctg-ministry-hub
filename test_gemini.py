import os
import time
from dotenv import load_dotenv
from google import genai

load_dotenv()

# Kunin ang dalawang API keys mula sa .env
api_keys = [
    os.getenv("GEMINI_API_KEY"),
    os.getenv("GEMINI_API_KEY_2")
]

# Salain ang mga walang laman
valid_keys = [k for k in api_keys if k and k.strip() != ""]

if not valid_keys:
    print("❌ Error: Walang nahanap na GEMINI_API_KEY o GEMINI_API_KEY_2 sa iyong .env file.")
else:
    print(f"🔍 Nakakita ng {len(valid_keys)} API key(s). Sinusubukan ang koneksyon...")
    
    success = False
    for idx, key in enumerate(valid_keys, start=1):
        print(f"\n🔑 Sinusubukan ang API Key #{idx}...")
        try:
            client = genai.Client(api_key=key)
            response = client.models.generate_content(
                model="gemini-3.8-flash",
                contents="Magbigay ng maikling pagbati sa Tagalog para sa isang Church Ministry Assistant."
            )
            print(f"✅ Tagumpay gamit ang API Key #{idx}!")
            print(f"🤖 Sagot ng AI: {response.text}")
            success = True
            break
        except Exception as e:
            print(f"⚠️ Nabigo ang API Key #{idx}: {str(e)}")
            continue
            
    if not success:
        print("\n❌ Nabigo ang lahat ng API keys. Subukan ulit mamaya.")