import time
import gspread
import os
import random
import re
import requests
import json
import sys
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from googleapiclient.discovery import build

# --- CONFIGURATION ---
SPREADSHEET_KEY = "********************************"
DOCUMENT_ID = "**************************"
INPUT_TAB_NAME = "Sheet1"
CREDENTIALS_FILE = "credentials.json"
HEADER_ROW_INDEX = 6 
GEMINI_API_KEY = "" # The environment provides this automatically

def get_google_creds():
    """Authenticates for Sheets, Drive, and Docs APIs."""
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/documents"
    ]
    if not os.path.exists(CREDENTIALS_FILE):
        raise FileNotFoundError(f"'{CREDENTIALS_FILE}' not found. Ensure the service account JSON is present.")
    return ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)

def setup_driver():
    """Sets up standard Selenium with advanced stealth flags."""
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    # Use a modern, common user agent
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36")
    
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    })
    
    return driver

def construct_review_url(url):
    """Extracts ASIN and builds a robust All-Reviews URL."""
    asin_match = re.search(r'/(?:dp|gp/product|product-reviews|product)/([A-Z0-9]{10})', url)
    if asin_match:
        asin = asin_match.group(1)
        domain_match = re.search(r'(https?://[a-zA-Z0-9.-]+\.[a-z]{2,})', url)
        domain = domain_match.group(1) if domain_match else "https://www.amazon.com"
        # Using a direct path to the customer reviews page is more stable than the product page
        return f"{domain}/product-reviews/{asin}/ref=cm_cr_arp_d_viewopt_srt?reviewerType=all_reviews&sortBy=recent&pageNumber=1"
    return url

def scrape_reviews_stealth(driver, url):
    """Navigates and scrapes reviews using multi-layer selectors."""
    driver.get(url)
    
    # 1. Wait for any common review container to appear
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div[data-hook='review'], .review, .a-section.review"))
        )
    except Exception:
        print("  -> Timeout: Reviews didn't load or page structure is different.")
    
    # 2. Human-like jitter
    time.sleep(random.uniform(2, 4))
    driver.execute_script("window.scrollTo(0, 1200);")
    time.sleep(1)
    
    soup = BeautifulSoup(driver.page_source, "html.parser")
    
    # Check for blocks using various selectors Amazon rotates through
    blocks = soup.find_all("div", {"data-hook": "review"}) or \
             soup.select(".a-section.review") or \
             soup.select(".review")
    
    parsed_reviews = []
    for block in blocks:
        try:
            # Extract Star Rating
            star_tag = block.find("i", {"data-hook": "review-star-rating"}) or \
                       block.find("i", {"data-hook": "cmps-review-star-rating"}) or \
                       block.select_one(".a-icon-star")
            
            score = 0
            if star_tag:
                score_str = star_tag.get_text() or star_tag.get('class', [''])[0]
                match = re.search(r'(\d)', score_str)
                score = int(match.group(1)) if match else 0

            # Extract Review Text
            text_tag = block.find("span", {"data-hook": "review-body"}) or \
                       block.select_one(".review-text") or \
                       block.select_one(".a-size-base.review-text")
            
            if text_tag:
                clean_text = text_tag.get_text(strip=True).replace("Read more", "").strip()
                if clean_text:
                    parsed_reviews.append({
                        "score": score,
                        "text": f"[{score} Stars] {clean_text[:500]}..."
                    })
        except Exception:
            continue
            
    return parsed_reviews

def summarize_with_ai(product_name, reviews):
    """Uses Gemini 1.5 Flash to summarize reviews."""
    if not reviews: return "No review data available."
    
    prompt = f"Analyze these reviews for '{product_name}'. Provide Sentiment, 3 Pros, 3 Cons, and Target Audience.\n\n{reviews}"
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={GEMINI_API_KEY}"
    payload = {"contents": [{"parts": [{"text": prompt}]}]}
    
    for attempt in range(3):
        try:
            res = requests.post(url, json=payload, timeout=15)
            return res.json()['candidates'][0]['content']['parts'][0]['text']
        except: time.sleep(2)
    return "AI Summary failed."

def append_to_doc(creds, doc_id, product_name, summary):
    """Updates Google Doc."""
    try:
        service = build('docs', 'v1', credentials=creds)
        body = f"\nPRODUCT: {product_name}\n" + "-"*20 + f"\n{summary}\n" + "="*40 + "\n"
        service.documents().batchUpdate(documentId=doc_id, body={'requests': [{'insertText': {'endOfSegmentLocation': {}, 'text': body}}]}).execute()
        print("  -> Google Doc updated.")
    except Exception as e: print(f"  -> Doc Error: {e}")

def main():
    print("--- Amazon Research Agent V3 (Advanced Scraper) ---")
    try:
        creds = get_google_creds()
        gc = gspread.authorize(creds)
        sheet = gc.open_by_key(SPREADSHEET_KEY).worksheet(INPUT_TAB_NAME)
        data = sheet.get_all_values()
        headers = data[HEADER_ROW_INDEX - 1]
        
        url_idx, name_idx = headers.index("URL"), headers.index("PRODUCT NAME")
        pos_col = headers.index("Most Positive Review") + 1 if "Most Positive Review" in headers else len(headers) + 1
        neg_col = headers.index("Most Negative Review") + 1 if "Most Negative Review" in headers else len(headers) + 2
    except Exception as e:
        print(f"Setup Error: {e}"); return

    driver = setup_driver()
    
    try:
        for i, row in enumerate(data[HEADER_ROW_INDEX:]):
            if not row[url_idx]: continue
            
            name = row[name_idx]
            print(f"\n[{i+1}] Scraping: {name}")
            
            review_url = construct_review_url(row[url_idx])
            reviews = scrape_reviews_stealth(driver, review_url)
            
            if reviews:
                # Update Sheet
                sorted_r = sorted(reviews, key=lambda x: x['score'], reverse=True)
                sheet.update_cell(HEADER_ROW_INDEX + 1 + i, pos_col, sorted_r[0]['text'])
                sheet.update_cell(HEADER_ROW_INDEX + 1 + i, neg_col, sorted_r[-1]['text'])
                print(f"  -> Found {len(reviews)} reviews. Sheet updated.")
                
                # Update Doc
                summary = summarize_with_ai(name, "\n".join([r['text'] for r in reviews[:10]]))
                append_to_doc(creds, DOCUMENT_ID, name, summary)
            else:
                print("  -> Failed to find reviews. Check if CAPTCHA appeared.")
            
            time.sleep(random.uniform(15, 25))
            
    finally:
        driver.quit()
        print("\n--- Session Finished ---")

if __name__ == "__main__":

    main()
