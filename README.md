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
This project is licensed under the MIT License - see the LICENSE file for details.
