import time
import gspread
import os
import random
import re
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup

# --- CONFIGURATION ---
SPREADSHEET_KEY = "1i-agaByyXpwkM0I7dUgGf9ZeUf0TVzsNykyadWyG8d0"
INPUT_TAB_NAME = "Sheet1"
CREDENTIALS_FILE = "credentials.json" 

def get_google_client():
    """Connects to Google Sheets API using your credentials.json file."""
    if not os.path.exists(CREDENTIALS_FILE):
        print(f"ERROR: '{CREDENTIALS_FILE}' not found.")
        print("Please download your Google Service Account key, rename it to 'credentials.json',")
        print("and place it in the same folder as this script.")
        raise FileNotFoundError("Credentials file missing.")
        
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)
    return client

def setup_local_driver():
    """Sets up Chrome for a local VS Code environment."""
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

def parse_rating(rating_text):
    try:
        if not rating_text:
            return 0.0
        return float(rating_text.split(" ")[0])
    except Exception:
        return 0.0

def construct_review_url(original_url):
    """Extracts the ASIN (Product ID) and builds a clean Review URL."""
    try:
        asin_match = re.search(r'/(?:dp|gp/product)/([A-Z0-9]{10})', original_url)
        if asin_match:
            asin = asin_match.group(1)
            domain_match = re.search(r'(https?://[^/]+)', original_url)
            domain = domain_match.group(1) if domain_match else "https://www.amazon.com"
            return f"{domain}/product-reviews/{asin}/ref=cm_cr_dp_d_show_all_btm?ie=UTF8&reviewerType=all_reviews"
        return original_url 
    except Exception as e:
        print(f"Error constructing URL: {e}")
        return original_url

def extract_reviews(driver, product_url):
    """Visits the product page, extracts review data, and returns the best and worst reviews."""
    
    # 1. Navigate
    target_url = construct_review_url(product_url)
    print(f"Navigating to: {target_url}")
    
    try:
        driver.get(target_url)
        
        # 2. CAPTCHA Check
        if "Robot Check" in driver.title or "CAPTCHA" in driver.title:
            print("  !!! ALERT: Amazon CAPTCHA detected. Please solve it manually !!!")
            print("  (Waiting 15 seconds...)")
            time.sleep(15)
        
        # 3. Scroll Logic
        time.sleep(3)
        print("  -> Scrolling page to load reviews...")
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        driver.execute_script("window.scrollTo(0, 1000);")
        time.sleep(1)

        soup = BeautifulSoup(driver.page_source, "html.parser")
        
        # 4. Search Strategies
        review_blocks = soup.find_all("div", {"data-hook": "review"})
        if not review_blocks:
            review_blocks = soup.find_all("div", id=lambda x: x and x.startswith('customer_review'))
        if not review_blocks:
            review_blocks = soup.find_all("div", class_="a-section review aok-relative")

        if not review_blocks:
            print("  ! No reviews found.")
            return None, None

        # 5. Parse Data
        print(f"  -> Found {len(review_blocks)} reviews. Analyzing sentiment...")
        reviews_data = []
        
        for block in review_blocks:
            try:
                # Get Rating
                rating_elem = block.find("i", {"data-hook": "review-star-rating"})
                if not rating_elem: rating_elem = block.find("i", class_="a-icon-star")
                if not rating_elem: rating_elem = block.find("span", class_="a-icon-alt")

                rating_text = rating_elem.text.strip() if rating_elem else "N/A"
                numeric_rating = parse_rating(rating_text)
                
                # Get Body
                body_elem = block.find("span", {"data-hook": "review-body"})
                body = body_elem.text.strip() if body_elem else ""
                
                # Limit body length to fit in cell nicely
                if len(body) > 500:
                    body = body[:500] + "..."

                if body:
                    reviews_data.append({
                        "NumericRating": numeric_rating,
                        "Review": f"[{rating_text}] {body}"
                    })
            except Exception:
                continue
        
        # 6. Find Best and Worst
        if reviews_data:
            sorted_reviews = sorted(reviews_data, key=lambda x: x['NumericRating'], reverse=True)
            best_review = sorted_reviews[0]['Review']
            worst_review = sorted_reviews[-1]['Review']
            return best_review, worst_review

        return None, None

    except Exception as e:
        print(f"  Error accessing URL: {e}")
        return None, None

def main():
    print("--- Amazon Review Agent Starting ---")
    
    try:
        client = get_google_client()
        input_sheet = client.open_by_key(SPREADSHEET_KEY).worksheet(INPUT_TAB_NAME)
        
        # Handle Custom Table Layout
        all_values = input_sheet.get_all_values()
        HEADER_ROW_INDEX = 5 # Row 6 in Sheet (0-based index is 5)
        
        if len(all_values) <= HEADER_ROW_INDEX:
            print("Error: Sheet is too empty.")
            return

        headers = all_values[HEADER_ROW_INDEX]
        
        # --- NEW: Automatically Add Columns if Missing ---
        
        # Check for 'Most Positive Review'
        if "Most Positive Review" not in headers:
            print("Creating 'Most Positive Review' column...")
            new_col_idx = len(headers) + 1
            input_sheet.update_cell(HEADER_ROW_INDEX + 1, new_col_idx, "Most Positive Review")
            headers.append("Most Positive Review") # Update local list
            
        # Check for 'Most Negative Review'
        if "Most Negative Review" not in headers:
            print("Creating 'Most Negative Review' column...")
            new_col_idx = len(headers) + 1
            input_sheet.update_cell(HEADER_ROW_INDEX + 1, new_col_idx, "Most Negative Review")
            headers.append("Most Negative Review") # Update local list

        # Find indexes (add 1 because gspread uses 1-based indexing)
        try:
            product_col_idx = headers.index("PRODUCT NAME")
            url_col_idx = headers.index("URL")
            pos_col_idx = headers.index("Most Positive Review") + 1
            neg_col_idx = headers.index("Most Negative Review") + 1
        except ValueError:
            print("CRITICAL ERROR: Could not find required columns in Row 6.")
            return

        print(f"Successfully connected. Columns ready. Processing rows...")
        
    except Exception as e:
        print(f"Connection Error: {e}")
        return

    print("Launching Browser...")
    driver = setup_local_driver()

    try:
        # We loop through data rows.
        # i starts at 0, which corresponds to the first data row.
        # The actual spreadsheet row number = HEADER_ROW_INDEX + 1 (header row) + 1 (first data) + i
        start_row_num = HEADER_ROW_INDEX + 2 
        
        # Slice the data to skip headers
        data_rows = all_values[HEADER_ROW_INDEX+1:]
        
        for i, row in enumerate(data_rows):
            current_row_num = start_row_num + i
            
            # Safety check for short rows
            if len(row) <= url_col_idx: 
                continue

            product_name = row[product_col_idx] # Python list index
            url = row[url_col_idx]              # Python list index

            if not url or "http" not in url:
                continue

            print(f"\nProcessing Row {current_row_num}: {product_name}")
            
            # Check if we already did this row (optional optimization)
            # If the positive review column is already filled, we can skip? 
            # Uncomment next 2 lines to skip already finished rows:
            # if len(row) >= pos_col_idx and row[pos_col_idx-1] != "":
            #     print("  -> Already processed. Skipping.")
            #     continue

            best_review, worst_review = extract_reviews(driver, url)
            
            if best_review:
                print("  -> Updating Sheet...")
                # Update specific cells in the SAME row
                input_sheet.update_cell(current_row_num, pos_col_idx, best_review)
                input_sheet.update_cell(current_row_num, neg_col_idx, worst_review)
                print("  -> Done!")
            else:
                print("  -> No reviews extracted.")
                input_sheet.update_cell(current_row_num, pos_col_idx, "No reviews found")
            
            print("Pausing for 5-10s...")
            time.sleep(random.uniform(5, 10))

    except KeyboardInterrupt:
        print("\nStopping Agent...")
    finally:
        driver.quit()
        print("\n--- Agent Finished ---")

if __name__ == "__main__":
    main()