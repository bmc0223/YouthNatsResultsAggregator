from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from collections import defaultdict
from urllib.parse import urlparse, parse_qs
import pandas as pd
import time
import os
import re
import pickle

# Define Desktop and cache path
desktop = os.path.join(os.path.expanduser("~"), "Desktop")
cache_file = os.path.join(desktop, "A_B_Finals_YN_2024_CACHE.pkl")

# Base event results page
MAIN_URL = "https://www.regattacentral.com/regatta/results2?job_id=8084&org_id=0"

# Initialize Selenium driver only if needed
driver = None
all_results = []
event_names = {}
refresh_cache = False  # Set to True if you want to force re-scraping

# Extract event_id from a URL
def get_event_id(url):
    return parse_qs(urlparse(url).query).get("event_id", ["?"])[0]

# Get all event URLs and names
def get_event_links():
    global driver
    driver.get(MAIN_URL)
    time.sleep(2)
    event_links = []
    event_names_local = {}
    anchors = driver.find_elements(By.CSS_SELECTOR, "a[href*='event_id=']")
    for a in anchors:
        href = a.get_attribute("href")
        text = a.text.strip()
        if "event_id=" in href:
            event_links.append(href)
            event_id = get_event_id(href)
            event_names_local[event_id] = text
    return list(set(event_links)), event_names_local

# Scrape results from each event URL
def get_finalists_and_times_selenium(url):
    global driver
    driver.get(url)
    time.sleep(2)
    results = []
    race_divs = driver.find_elements(By.CSS_SELECTOR, "div[id^='Race']")
    for race_div in race_divs:
        try:
            header_table = race_div.find_element(By.CLASS_NAME, "raceHeader")
            header_text = header_table.text
            if "Final A" not in header_text and "Final B" not in header_text:
                continue
            final_type = "Final A" if "Final A" in header_text else "Final B"
            results_table = race_div.find_element(By.CSS_SELECTOR, "table.tablesorter")
            rows = results_table.find_elements(By.TAG_NAME, "tr")[1:]
            for row in rows:
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) < 6:
                    continue
                place = cols[0].text.strip()
                org = cols[3].text.strip()
                finish_time = cols[5].text.strip().split("\n")[0]
                results.append({
                    "final": final_type,
                    "place": place,
                    "organization": org,
                    "finish_time": finish_time,
                    "event_url": url,
                })
        except Exception as e:
            print(f"Error parsing race div in {url}: {e}")
            continue
    return results

# Load from cache or scrape fresh
def load_or_scrape():
    global driver, event_names, all_results

    if not refresh_cache and os.path.exists(cache_file):
        print("\n🗃️  Loading cached results...")
        with open(cache_file, "rb") as f:
            all_results = pickle.load(f)
        return

    # Set up headless Selenium driver
    options = Options()
    options.add_argument("--headless")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    print("Fetching event links...")
    event_urls, event_names_local = get_event_links()
    event_names.update(event_names_local)
    print(f"Found {len(event_urls)} events.\n")

    for i, event_url in enumerate(event_urls):
        print(f"[{i+1}/{len(event_urls)}] Processing {event_url}")
        results = get_finalists_and_times_selenium(event_url)
        all_results.extend(results)

    driver.quit()

    with open(cache_file, "wb") as f:
        pickle.dump(all_results, f)
    print(f"\n📦 Cached results saved to {cache_file}")

# Run
load_or_scrape()

# Group results by event ID
grouped_results = defaultdict(list)
for entry in all_results:
    event_id = get_event_id(entry.get("event_url", "unknown"))
    grouped_results[event_id].append(entry)

# Save to Excel
output_file = os.path.join(desktop, "A_B_Finals_YN_2024.xlsx")
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for event_id, results in grouped_results.items():
        df = pd.DataFrame(results)
        columns = ["final", "place", "organization", "finish_time"]
        df = df[[col for col in columns if col in df.columns]]

        # Sort so Final A rows appear before Final B
        df["final_sort"] = df["final"].apply(lambda x: 0 if x == "Final A" else 1)
        df = df.sort_values(by=["final_sort", "place"])
        df = df.drop(columns=["final_sort"])

        raw_name = event_names.get(event_id, f"Event_{event_id}")
        clean_name = re.sub(r'[:\\/?*\[\]]', '-', raw_name)
        sheet_name = clean_name[:31]

        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"\n✅ Excel file with one sheet per event saved to:\n{output_file}")
