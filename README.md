# 🚛 Tamil Nadu — Used Commercial Vehicle Dealerships Scraper

A Python web scraper that collects business data for **Used Commercial Vehicle Dealerships** across all **38 districts of Tamil Nadu** from Google Maps — with built-in anti-detection, auto-save, and multi-format export.

---

## 📌 What It Scrapes

Targets dealerships selling used trucks, lorries, buses, tippers, JCBs, and tractors across Tamil Nadu.

**Data collected per dealership:**

| Field | Example |
|---|---|
| District | Chennai |
| Name | Sri Murugan Used Trucks |
| Address | NH-48, Bypass Road, Ambattur |
| Phone | +91 98765 43210 |
| Rating | 4.3 |
| Total Reviews | 245 |
| Google Maps URL | https://maps.google.com/... |

---

## 🏗️ Tech Stack

| Tool | Purpose |
|---|---|
| `playwright` | Browser automation (Chromium) |
| `pandas` | Data processing & export |
| `openpyxl` | Excel file writing |

> Uses **Playwright** instead of Selenium — faster, more reliable, and better at handling Google Maps' dynamic content.

---

## 📦 Installation

**Prerequisites:** Python 3.8+

```bash
pip install playwright pandas openpyxl
playwright install chromium
```

> The second command downloads the Chromium browser that Playwright uses — run it once after install.

---

## 🚀 Usage

```bash
python commercial_vehicle_scraper.py
```

The scraper loops through all 38 districts, runs 5 search queries per district, deduplicates results, auto-saves progress after each district, and writes final output files when done.

---

## ⚙️ Configuration

All settings are at the top of the script — no need to touch anything else:

```python
HEADLESS = True
# True  → browser runs silently in background (faster)
# False → watch the browser live (great for debugging)

SCROLL_COUNT = 12
# More scrolls = more results per query (recommended: 10–15)

DELAY_BETWEEN_SEARCHES  = (3, 6)    # seconds between queries (random range)
DELAY_BETWEEN_DISTRICTS = (8, 14)   # seconds between districts (random range)
```

---

## 🔍 Search Queries (per district)

The scraper runs **5 different queries** per district to maximize coverage:

```
1. used commercial vehicle dealership {district} Tamil Nadu
2. second hand lorry truck dealer {district} Tamil Nadu
3. used trucks buses dealership {district}
4. second hand commercial vehicle {district} Tamil Nadu
5. used tipper JCB tractor dealer {district} Tamil Nadu
```

Duplicates found across queries are automatically removed.

---

## 📁 Output Files

| File | Description |
|---|---|
| `used_commercial_vehicles.csv` | All results in one CSV (UTF-8, Excel-compatible) |
| `used_commercial_vehicles.xlsx` | Excel with **one sheet per district** + an "ALL DISTRICTS" sheet |
| `progress_backup.csv` | Auto-saved after every district — protects your data if the script crashes |

---

## 🛡️ Anti-Detection Features

- Randomized delays between queries and districts
- Realistic User-Agent string (mimics Chrome 124 on Windows)
- `navigator.webdriver` property hidden via `add_init_script`
- `--disable-blink-features=AutomationControlled` Chrome flag
- Human-like viewport and locale settings

---

## 📊 Sample Summary Output

```
============================================================
  🚛 Tamil Nadu — Used Commercial Vehicle Dealerships
  📌 Districts to scrape : 38
  🔍 Searches per district: 5
============================================================

──────────────────────────────────────────────────────
  [01/38] 📍 District: Ariyalur
──────────────────────────────────────────────────────
  ✓ 18 dealerships found in Ariyalur
  💾 Progress saved (18 total so far)

  ...

──────────────────────────────────────────────────────
  📊 Total dealerships found : 612
  📍 Districts covered       : 38 / 38
  ⭐ Average rating          : 3.87
──────────────────────────────────────────────────────
```

---

## 🧯 Troubleshooting

**Getting no results?** Try these fixes:

1. Set `HEADLESS = False` to watch the browser and spot issues
2. Increase delays: `DELAY_BETWEEN_SEARCHES = (8, 12)`
3. Check your internet connection
4. Google Maps may have updated its HTML — open an issue if selectors break

---

## ⚠️ Disclaimer

This project is for **educational and research purposes only**. Scraping Google Maps may violate [Google's Terms of Service](https://policies.google.com/terms). Use responsibly and ensure compliance with applicable laws before using this in any production or commercial context.

---

## 🤝 Contributing

Pull requests are welcome! If Google Maps updates its CSS class names and the scraper breaks, feel free to open an issue or submit a fix.

---

## 📄 License

MIT License — free to use, modify, and distribute.
