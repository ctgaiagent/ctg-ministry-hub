import os
import io
import requests
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

# Kunin ang mga API keys mula sa .env o environment
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
NOTION_TOKEN = os.getenv("NOTION_TOKEN")
NOTION_CRM_DB_ID = os.getenv("NOTION_CRM_DB_ID", "409fff6803d78068bdaffcb10ce088ab")

CLOUD_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/15_znwPmjK95vTgf5lStbKKPCjEGw6oDp/edit?usp=sharing&ouid=107511419052324683693&rtpof=true&sd=true"

PARISH_ORDER = [
	"LIPAY NORTE PARISH", "MALANAS PARISH", "POBLACION PARISH", "DUPLAS PARISH",
	"SANTIAGO PARISH", "CABA PARISH", "MANILA PARISH", "ANDABUEN PARISH",
	"APAYAO PARISH", "HONGKONG PARISH"
]

def extract_id_from_url(url: str) -> str:
	import re
	match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
	return match.group(1) if match else url

def test_google_drive_excel():
	print("----------------------------------------")
	print("🔍 Sinusubukan ang Google Drive Excel...")
	doc_id = extract_id_from_url(CLOUD_SPREADSHEET_URL)
	download_url = f"https://docs.google.com/spreadsheets/d/{doc_id}/export?format=xlsx"
    
	try:
		res = requests.get(download_url, timeout=30)
		if res.status_code == 200 and len(res.content) > 1000:
			xls = pd.ExcelFile(io.BytesIO(res.content))
			print(f"✅ Tagumpay! Nabuksan ang Excel. Mga Sheets: {xls.sheet_names}")
		else:
			print(f"❌ Nabigo sa pag-download. Status Code: {res.status_code}")
	except Exception as e:
		print(f"❌ May error sa pagbasa ng Excel: {str(e)}")

def test_notion_crm():
	print("----------------------------------------")
	print("🔍 Sinusubukan ang Notion CRM Database...")
	if not NOTION_TOKEN:
		print("⚠️ Babala: Walang nahanap na NOTION_TOKEN sa .env file.")
		return
        
	clean_id = NOTION_CRM_DB_ID.strip().split("?")[0].split("-")[-1]
	url = f"https://api.notion.com/v1/databases/{clean_id}/query"
	headers = {
		"Authorization": f"Bearer {NOTION_TOKEN.strip()}",
		"Notion-Version": "2022-06-28",
		"Content-Type": "application/json"
	}
    
	try:
		res = requests.post(url, headers=headers, json={"page_size": 10}, timeout=20)
		if res.status_code == 200:
			data = res.json()
			results = data.get("results", [])
			print(f"✅ Tagumpay! Nakakonekta sa Notion CRM. May nakitang {len(results)} na bisita/rekord.")
			for page in results:
				props = page.get("properties", {})
				name_prop = props.get("Full Name", {}).get("title", [])
				name = name_prop[0].get("plain_text", "Wala pang pangalan") if name_prop else "Wala"
				print(f"   - Bisita: {name}")
		else:
			print(f"❌ Nabigo ang Notion API. Status Code: {res.status_code} - Responde: {res.text}")
	except Exception as e:
		print(f"❌ May error sa Notion connection: {str(e)}")

if __name__ == "__main__":
	test_google_drive_excel()
	test_notion_crm()
	print("----------------------------------------")
	print("✨ Tapos na ang test!")
