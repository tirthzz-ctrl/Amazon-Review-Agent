SAMPLE DATA USED - https://docs.google.com/spreadsheets/d/1i-agaByyXpwkM0I7dUgGf9ZeUf0TVzsNykyadWyG8d0/edit?gid=0#gid=0


# 🛒 Amazon Review Analyzer - AI Enhanced

This is a specialized AI workflow that scrapes Amazon product reviews using a **Web Automation Agent**, ranks them using **Claude AI** based on strategic metrics, and automatically syncs the results back to **Google Sheets**.

## 🎯 Core Features

* **Automated Scraping**: Uses a headless Web Agent to navigate Amazon product pages and extract customer feedback.
* **AI Ranking & Analysis**: Leverages `claude-haiku-4-5` to score reviews on **Helpfulness**, **Persuasiveness**, and **Authenticity**.
* **Intelligent Recommendations**: AI generates strategic advice on how to use specific reviews for marketing or product improvement.
* **Google Sheets Integration**: Updates three distinct columns (Positive, Negative, and Recommendations) automatically.

---

## 🚀 Quick Start (Windows Setup)

Since you are running this on Windows, follow these exact steps to avoid environment errors.

### 1. Environment Setup

```powershell
# Create the virtual environment
python -m venv .venv

# Activate the environment
.\.venv\Scripts\Activate.ps1

# Install required libraries
pip install codewords-client==0.4.0 fastapi==0.116.1 anthropic==0.62.0 uvicorn[standard]

```

### 2. Configuration (`.env`)

Create a file named `.env` in your root directory and add your credentials:

```env
PORT=8000
LOGLEVEL=INFO
CODEWORDS_API_KEY=your_key_here
CODEWORDS_RUNTIME_URI=https://runtime.codewords.ai
ANTHROPIC_API_KEY=your_claude_key_here
PIPEDREAM_GOOGLE_SHEETS_ACCESS=your_token_here

```

---

## 🛠 Critical Windows Fix

The script includes a built-in fix for the `uvloop` incompatibility issue common on Windows. Ensure this block remains at the top of your `main.py`:

```python
import asyncio
import sys

if sys.platform == 'win32':
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

```

---

## 📊 Data Mapping

The service processes Google Sheets data using the following default column indices (0-based):

| Column | Index | Content |
| --- | --- | --- |
| **Column J** | 9 | **Input**: Amazon Product URL |
| **Column K** | 10 | **Output**: Top 5 Positive Reviews (Ranked) |
| **Column L** | 11 | **Output**: Top 5 Negative Reviews (Ranked) |
| **Column M** | 12 | **Output**: AI Strategic Recommendations |

---

## 🖥 Running the Service

1. **Start the Server**:
```powershell
python main.py

```


Amazon Research Agent V3 (Advanced Scraper)

An automated research tool that scrapes Amazon product reviews using Selenium, analyzes customer sentiment using Google Gemini 1.5 Flash AI, and synchronizes the results directly into Google Sheets and Google Docs.

🚀 Features

Stealth Scraping: Utilizes advanced Selenium configurations and randomized human-like behavior to bypass basic bot detection.

Dynamic Parsing: Multi-layer selectors to handle Amazon's rotating HTML structure and dynamic content loading.

AI-Powered Analysis: Integrates with Gemini 1.5 Flash to generate professional summaries including sentiment, pros/cons, and target audience recommendations.

Bi-Directional Sync:

Google Sheets: Automatically identifies and populates the "Most Positive" and "Most Negative" reviews for specific products.

Google Docs: Appends structured research reports for each product investigated.

Robust Error Handling: Includes detection checks for CAPTCHAs and "Robot Checks" with manual bypass support.

📋 Prerequisites

Before running the agent, ensure you have the following:

Python 3.10+

Google Cloud Service Account: A JSON credentials file with access to:

Google Sheets API

Google Drive API

Google Docs API

Gemini API Key: Obtainable via Google AI Studio.

Chrome Browser: Installed on the host machine (compatible with ChromeDriver).

🛠️ Installation

Clone the repository:

git clone [https://github.com/yourusername/amazon-research-agent.git]
(https://github.com/yourusername/amazon-research-agent.git)
cd amazon-research-agent


Install dependencies:

pip install -r requirements.txt


Required packages: gspread, oauth2client, selenium, webdriver-manager, beautifulsoup4, google-api-python-client, requests.

Setup Credentials:

Place your Service Account JSON file in the project root and name it credentials.json.

Share your target Google Sheet and Google Doc with the email address found in your credentials.json.

⚙️ Configuration

Open amazon_research_agent.py and update the following constants:

SPREADSHEET_KEY = "your_google_sheet_id_here"
DOCUMENT_ID = "your_google_doc_id_here"
INPUT_TAB_NAME = "Sheet1"
HEADER_ROW_INDEX = 6  # Row number where your headers (URL, PRODUCT NAME) are located


The GEMINI_API_KEY is typically provided via environment variables or direct input in the script.

📖 Usage

Run the agent from your terminal:

python amazon_research_agent.py


Process Flow:

The agent reads the list of URLs and Product Names from your Google Sheet starting at HEADER_ROW_INDEX.

It opens a stealth Chrome instance and navigates to the "All Reviews" page for the product.

It scrolls and mimics human behavior to load reviews.

If a CAPTCHA is detected, the script will pause. Solve it manually in the browser window to continue.

Once parsed, it updates the Google Sheet with the best/worst reviews.

Finally, it generates an AI summary and appends it to your Google Doc.

⚠️ Important Notes

Anti-Ban Strategy: The agent includes random cooldowns (15-25 seconds) between products. Do not shorten these significantly, or you risk an IP ban from Amazon.

Sheet Structure: Ensure your Google Sheet contains headers named exactly URL, PRODUCT NAME, Most Positive Review, and Most Negative Review.

Headless Mode: This script is configured to run with a visible browser window to allow for manual CAPTCHA solving if triggered.

📄 License

This project is licensed under the MIT License - see the LICENSE file for details.
