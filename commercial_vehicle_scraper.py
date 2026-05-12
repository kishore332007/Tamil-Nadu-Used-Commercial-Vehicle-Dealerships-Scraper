"""
╔══════════════════════════════════════════════════════════════════╗
║   GOOGLE MAPS SCRAPER — Tamil Nadu                               ║
║   Used Commercial Vehicle Dealerships                            ║
║   Scrapes all 38 TN districts                                    ║
║   Data: Name, Phone, Address, Rating, Reviews                    ║
╚══════════════════════════════════════════════════════════════════╝

STEP 1 — Install required libraries (run once):
    pip install playwright pandas openpyxl
    playwright install chromium

STEP 2 — Run the scraper:
    python scraper.py

OUTPUT FILES:
    used_commercial_vehicles.csv   → CSV file (opens in Excel / Google Sheets)
    used_commercial_vehicles.xlsx  → Excel file (one sheet per district)
    progress_backup.csv            → Auto-saved after every district
"""

# ─────────────────────────────────────────────────────────────────
#  IMPORTS — Libraries we need
# ─────────────────────────────────────────────────────────────────

import time        # Used for sleep/delay between requests
import random      # Used for random delays so we don't look like a bot
import re          # Used for detecting phone numbers using pattern matching
import pandas as pd                          # Used for saving Excel/CSV files
from playwright.sync_api import sync_playwright  # Browser automation library


# ─────────────────────────────────────────────────────────────────
#  SETTINGS — Change these values if needed
# ─────────────────────────────────────────────────────────────────

HEADLESS = True
# True  → browser runs silently in background (faster)
# False → you can watch the browser scrape live (good for debugging)

SCROLL_COUNT = 12
# How many times to scroll down in results panel
# More scrolls = more results loaded, but takes longer
# Recommended: 10–15

DELAY_BETWEEN_SEARCHES  = (3, 6)
# Random wait (in seconds) between each search query
# Keeps us from being blocked by Google

DELAY_BETWEEN_DISTRICTS = (8, 14)
# Random wait (in seconds) between districts
# Slightly longer to avoid detection

OUTPUT_CSV  = "used_commercial_vehicles.csv"
OUTPUT_XLSX = "used_commercial_vehicles.xlsx"


# ─────────────────────────────────────────────────────────────────
#  ALL 38 TAMIL NADU DISTRICTS
# ─────────────────────────────────────────────────────────────────

TN_DISTRICTS = [
    "Ariyalur",        "Chengalpattu",    "Chennai",         "Coimbatore",
    "Cuddalore",       "Dharmapuri",      "Dindigul",        "Erode",
    "Kallakurichi",    "Kancheepuram",    "Kanyakumari",     "Karur",
    "Krishnagiri",     "Madurai",         "Mayiladuthurai",  "Nagapattinam",
    "Namakkal",        "Nilgiris",        "Perambalur",      "Pudukkottai",
    "Ramanathapuram",  "Ranipet",         "Salem",           "Sivaganga",
    "Tenkasi",         "Thanjavur",       "Theni",           "Thoothukudi",
    "Tiruchirappalli", "Tirunelveli",     "Tirupathur",      "Tiruppur",
    "Tiruvallur",      "Tiruvannamalai",  "Tiruvarur",       "Vellore",
    "Viluppuram",      "Virudhunagar",
]


# ─────────────────────────────────────────────────────────────────
#  SEARCH QUERIES
#  We run multiple searches per district to get maximum results
# ─────────────────────────────────────────────────────────────────

SEARCH_QUERIES = [
    "used commercial vehicle dealership {district} Tamil Nadu",
    "second hand lorry truck dealer {district} Tamil Nadu",
    "used trucks buses dealership {district}",
    "second hand commercial vehicle {district} Tamil Nadu",
    "used tipper JCB tractor dealer {district} Tamil Nadu",
]


# ─────────────────────────────────────────────────────────────────
#  FUNCTION 1: extract_from_card()
#  PURPOSE: Read data from ONE result card on Google Maps
# ─────────────────────────────────────────────────────────────────

def extract_from_card(card, district):
    """
    INPUT : A single Google Maps result card (HTML element)
    OUTPUT: A dictionary with name, phone, address, rating, reviews

    Google Maps shows each business as a "card" in the left panel.
    This function reads one card and extracts all the info we need.
    """

    # Start with empty values — we'll fill them in below
    data = {
        "district":      district,
        "name":          "",
        "address":       "",
        "phone":         "",
        "rating":        "",
        "total_reviews": "",
        "maps_url":      "",
    }

    try:
        # ── EXTRACT NAME ──────────────────────────────────────────────────
        # The business name is in a div with CSS class "fontHeadlineSmall"
        # Example: "Sri Murugan Used Trucks"
        name_el = card.query_selector("div.fontHeadlineSmall")
        if name_el:
            data["name"] = name_el.inner_text().strip()

        # If we can't find a name, skip this card entirely
        if not data["name"]:
            return None

        # ── EXTRACT RATING ────────────────────────────────────────────────
        # Star rating is in a span with class "MW4etd"
        # Example: "4.3"
        rating_el = card.query_selector("span.MW4etd")
        if rating_el:
            data["rating"] = rating_el.inner_text().strip()

        # ── EXTRACT TOTAL REVIEWS ─────────────────────────────────────────
        # Review count is in a span with class "UY7F9"
        # Example: "(245)" → we strip brackets to get "245"
        reviews_el = card.query_selector("span.UY7F9")
        if reviews_el:
            raw = reviews_el.inner_text().strip()
            data["total_reviews"] = raw.strip("()")  # Remove ( and )

        # ── EXTRACT ADDRESS & PHONE ───────────────────────────────────────
        # Both address and phone appear inside ".W4Efsd" spans
        # We check each span's text to decide if it's a phone or address

        all_spans = card.query_selector_all("div.W4Efsd span")

        for span in all_spans:
            text = span.inner_text().strip()

            if not text or len(text) < 3:
                continue   # Skip empty or very short text

            # PHONE CHECK: Starts with digit or +, contains mostly numbers
            # Examples: "9876543210", "+91 98765 43210", "044-23456789"
            is_phone = re.match(r'^[\+\d][\d\s\(\)\-]{7,}$', text)
            if is_phone:
                data["phone"] = text
                continue

            # ADDRESS CHECK: Contains words common in Indian addresses
            address_keywords = [
                "street", "road", "nagar", "salai", "main", "cross",
                "bypass", "highway", "nh", "kovil", "market", "bus stand",
                "no.", "#", "th", "st", "nd", "district", "tamil",
                "near", "opp", "opposite", "behind", "floor"
            ]
            text_lower = text.lower()
            is_address = any(word in text_lower for word in address_keywords)

            if is_address and not data["address"] and len(text) > 5:
                data["address"] = text

        # ── EXTRACT GOOGLE MAPS LINK ──────────────────────────────────────
        # Each card has an <a> tag with the full Google Maps URL
        link_el = card.query_selector("a[href*='google.com/maps']")
        if link_el:
            data["maps_url"] = link_el.get_attribute("href") or ""

    except Exception as e:
        # If anything goes wrong reading this card, skip it
        print(f"        ⚠ Error reading card: {e}")
        return None

    return data


# ─────────────────────────────────────────────────────────────────
#  FUNCTION 2: scrape_one_search()
#  PURPOSE: Open Google Maps for ONE search query and collect results
# ─────────────────────────────────────────────────────────────────

def scrape_one_search(page, query, district):
    """
    INPUT : Playwright page object, search query string, district name
    OUTPUT: List of dealership dictionaries found for that query

    What this function does:
    1. Opens Google Maps with the search query
    2. Scrolls down multiple times to load more results
    3. Finds all result cards and extracts data from each
    """

    results    = []        # Will store all found dealerships
    seen_names = set()     # Tracks names already added (to avoid duplicates)

    # Build the Google Maps URL
    # Example: https://www.google.com/maps/search/used+commercial+vehicle+Chennai
    url = "https://www.google.com/maps/search/" + query.replace(" ", "+")
    print(f"      🔍 {query}")

    try:
        # ── Open the Google Maps search page ─────────────────────────────
        page.goto(url, timeout=30000, wait_until="domcontentloaded")
        time.sleep(3)  # Give the page time to fully load

        # ── Dismiss cookie consent popup (appears sometimes) ──────────────
        try:
            page.click("button:has-text('Accept all')", timeout=3000)
            time.sleep(1)
        except Exception:
            pass  # No popup found, that's fine — continue

        # ── Wait for result cards to appear ───────────────────────────────
        # Result cards have CSS class "Nv2PK"
        try:
            page.wait_for_selector("div.Nv2PK", timeout=10000)
        except Exception:
            print(f"      ⚠ No results loaded for: {query}")
            return results  # Return empty list if no results

        # ── Scroll down to load more results ──────────────────────────────
        # Google Maps loads results lazily — we must scroll to see more
        results_panel = page.query_selector("div[role='feed']")  # Left panel

        for scroll_num in range(1, SCROLL_COUNT + 1):

            # Scroll the left results panel down by 800px
            if results_panel:
                results_panel.evaluate("el => el.scrollTop += 800")
            else:
                page.evaluate("window.scrollBy(0, 800)")  # Fallback

            time.sleep(1.5)  # Wait for new results to load after scroll

            # Check if we've reached the end of all results
            end_markers = page.query_selector_all("span, p")
            for marker in end_markers:
                try:
                    if "end of the list" in marker.inner_text().lower():
                        print(f"      ✓ Reached end at scroll {scroll_num}")
                        break
                except Exception:
                    pass

        # ── Collect all result cards from the page ────────────────────────
        cards = page.query_selector_all("div.Nv2PK")
        print(f"      📋 {len(cards)} cards found")

        # ── Extract data from each card ───────────────────────────────────
        for card in cards:
            dealership = extract_from_card(card, district)

            # Skip if extraction failed
            if not dealership or not dealership["name"]:
                continue

            # Skip duplicates
            name_key = dealership["name"].lower().strip()
            if name_key in seen_names:
                continue

            seen_names.add(name_key)
            results.append(dealership)

    except Exception as e:
        print(f"      ❌ Page error: {e}")

    return results


# ─────────────────────────────────────────────────────────────────
#  FUNCTION 3: scrape_district()
#  PURPOSE: Run ALL search queries for ONE district
# ─────────────────────────────────────────────────────────────────

def scrape_district(page, district):
    """
    INPUT : Playwright page object, district name
    OUTPUT: List of all unique dealerships found in that district

    Runs 5 different search queries for the district.
    Combines all results and removes duplicates across queries.
    """

    district_results = []    # All dealerships for this district
    district_names   = set() # Names already added (for dedup across queries)

    for query_template in SEARCH_QUERIES:

        # Fill in the district name into the query template
        # Example: "used commercial vehicle {district}" → "used commercial vehicle Chennai"
        query = query_template.replace("{district}", district)

        # Scrape this one query
        found = scrape_one_search(page, query, district)

        # Add only NEW dealerships (not seen in previous queries)
        for item in found:
            name_key = item["name"].lower().strip()
            if name_key not in district_names:
                district_names.add(name_key)
                district_results.append(item)

        # Wait before next query to avoid being blocked
        delay = random.uniform(*DELAY_BETWEEN_SEARCHES)
        print(f"      ⏳ Waiting {delay:.1f}s...")
        time.sleep(delay)

    return district_results


# ─────────────────────────────────────────────────────────────────
#  FUNCTION 4: save_results()
#  PURPOSE: Save all collected data to CSV and Excel files
# ─────────────────────────────────────────────────────────────────

def save_results(all_data):
    """
    INPUT : List of all dealership dictionaries
    OUTPUT: Saves CSV and Excel files to disk

    Excel file has:
    - Sheet 1: All districts combined
    - One extra sheet per district
    """

    columns = ["district", "name", "address", "phone",
               "rating", "total_reviews", "maps_url"]

    # Create DataFrame
    df = pd.DataFrame(all_data, columns=columns)

    # Remove exact duplicates
    df.drop_duplicates(subset=["name", "district"], inplace=True)

    # Convert rating to number so we can sort (some ratings might be blank)
    df["rating"] = pd.to_numeric(df["rating"], errors="coerce")

    # Sort: by district A→Z, then by rating high→low
    df.sort_values(["district", "rating"], ascending=[True, False], inplace=True)
    df.reset_index(drop=True, inplace=True)

    # ── Save CSV ──────────────────────────────────────────────────────────
    # utf-8-sig encoding ensures Tamil characters display correctly in Excel
    df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
    print(f"\n  ✅ CSV  saved → {OUTPUT_CSV}")

    # ── Save Excel (multiple sheets) ──────────────────────────────────────
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:

        # First sheet: all districts together
        df.to_excel(writer, sheet_name="ALL DISTRICTS", index=False)

        # One sheet per district
        for district_name, group in df.groupby("district"):
            sheet_name = district_name[:31]  # Excel max sheet name = 31 chars
            group.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"  ✅ Excel saved → {OUTPUT_XLSX}")

    # Print summary
    print(f"\n  {'─'*45}")
    print(f"  📊 Total dealerships found : {len(df)}")
    print(f"  📍 Districts covered       : {df['district'].nunique()} / {len(TN_DISTRICTS)}")
    print(f"  ⭐ Average rating          : {df['rating'].mean():.2f}")
    print(f"  {'─'*45}")

    # Show top 5 results as preview
    print(f"\n  Top 5 Results Preview:")
    print(df[["district", "name", "rating"]].head(5).to_string(index=False))


# ─────────────────────────────────────────────────────────────────
#  MAIN — This is where everything starts
# ─────────────────────────────────────────────────────────────────

def main():
    print("\n" + "=" * 60)
    print("  🚛 Tamil Nadu — Used Commercial Vehicle Dealerships")
    print(f"  📌 Districts to scrape : {len(TN_DISTRICTS)}")
    print(f"  🔍 Searches per district: {len(SEARCH_QUERIES)}")
    print(f"  📁 CSV Output  : {OUTPUT_CSV}")
    print(f"  📁 Excel Output: {OUTPUT_XLSX}")
    print("=" * 60 + "\n")

    all_data = []  # Master list — stores every dealership from every district

    # ── Launch the browser ────────────────────────────────────────────────
    with sync_playwright() as p:

        # Launch Chromium browser
        browser = p.chromium.launch(
            headless=HEADLESS,
            args=[
                "--no-sandbox",
                "--disable-blink-features=AutomationControlled",
                # ↑ Hides the fact that this is an automated browser
            ]
        )

        # Create a browser tab that looks like a real human user
        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            locale="en-US",
            viewport={"width": 1280, "height": 900},
        )

        page = context.new_page()

        # Remove the "webdriver" JavaScript property
        # Websites check this to detect bots — we hide it
        page.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', { get: () => undefined })"
        )

        # ── Loop through all 38 districts ─────────────────────────────────
        for index, district in enumerate(TN_DISTRICTS, start=1):

            print(f"\n{'─'*55}")
            print(f"  [{index:02}/{len(TN_DISTRICTS)}] 📍 District: {district}")
            print(f"{'─'*55}")

            # Scrape all search queries for this district
            district_data = scrape_district(page, district)

            if district_data:
                all_data.extend(district_data)
                print(f"\n  ✓ {len(district_data)} dealerships found in {district}")
            else:
                print(f"\n  ✗ No results found for {district}")

            # ── Auto-save progress after every district ────────────────────
            # This protects your data if the script crashes halfway
            if all_data:
                pd.DataFrame(all_data).to_csv(
                    "progress_backup.csv",
                    index=False,
                    encoding="utf-8-sig"
                )
                print(f"  💾 Progress saved ({len(all_data)} total so far)")

            # Wait before next district (skip wait after last district)
            if index < len(TN_DISTRICTS):
                pause = random.uniform(*DELAY_BETWEEN_DISTRICTS)
                print(f"  ⏳ Waiting {pause:.0f}s before next district...")
                time.sleep(pause)

        browser.close()
        print("\n  🔒 Browser closed")

    # ── Save final output files ────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("  💾 Saving final output files...")
    print("=" * 60)

    if all_data:
        save_results(all_data)
    else:
        print("\n  ⚠ No data collected at all.")
        print("  Try:")
        print("    1. Set HEADLESS = False to watch the browser")
        print("    2. Increase DELAY_BETWEEN_SEARCHES to (8, 12)")
        print("    3. Check your internet connection")

    print("\n  🏁 Scraping complete!\n")


# ── Script entry point ─────────────────────────────────────────────────────────
# This block only runs when you execute: python scraper.py
# (Not when this file is imported by another script)

if __name__ == "__main__":
    main()
