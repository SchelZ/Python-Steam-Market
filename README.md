# 🎮 Steam Price Scraper & GUI Analyzer

A Tkinter-based desktop application that retrieves the latest Steam game listings, extracts **names & prices**, and allows you to:

* 📥 Retrieve live data from Steam (web scraping)
* 📊 Generate a price bar chart (Canvas-based)
* 📋 Display the full price/name matrix
* 📁 Export results to a **formatted Excel file** (`.xlsx`)

---

## 🚀 Features

* Uses **BeautifulSoup + urllib** to fetch Steam data
* Detects **Free** vs **Paid** games automatically
* Filters games with name length < 10 chars for graph clarity
* Bar chart shows **name + price** (with "Free" support)
* Excel export with **auto column width** formatting

---

## ✅ Dependencies

These must be installed manually:

```bash
pip install beautifulsoup4 pandas openpyxl
```

✅ `tkinter`, `urllib`, `typing`, `collections` → already included in Python (no install needed)

---

## ▶️ Run

```bash
python main.py
```

---
