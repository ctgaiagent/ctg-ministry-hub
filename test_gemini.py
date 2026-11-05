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
    max_retries = 3
    
    for idx, key in enumerate(valid_keys, start=1):
        print(f"\n🔑 Sinusubukan ang API Key #{idx}...")
        
        key_success = False
        for attempt in range(max_retries):
            try:
                print(f"   - Pagsubok #{attempt + 1}...")
                client = genai.Client(api_key=key)
                response = client.models.generate_content(
                    model="gemini-3.8-flash",
                    contents="Magbigay ng maikling pagbati sa Tagalog para sa isang Church Ministry Assistant."
                )
                print(f"✅ Tagumpay gamit ang API Key #{idx}!")
                print(f"🤖 Sagot ng AI: {response.text}")
                key_success = True
                success = True
                break
            except Exception as sub_err:
                print(f"   ⚠️ Pagsubok #{attempt + 1} nabigo: {str(sub_err)}")
                if attempt < max_retries - 1:
                    time.sleep(3)  # Maghintay ng 3 segundo bago ulitin
                else:
                    print(f"   ❌ Naubos ang pagsubok para sa API Key #{idx}.")
                    
        if key_success:
            break
            
    if not success:
        print("\n❌ Kasalukuyang nakakaranas ng mataas na traffic (503) ang mga server ng Gemini. Subukan itong patakbuhin mamaya pagkalipas ng ilang minuto.")