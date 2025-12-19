import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time
import re

# Path to Excel file
excel_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "ScrappedProducts.xlsx")

# Read Excel file
df = pd.read_excel(excel_path)

# Get first product
first_row = df.iloc[0]
link = first_row["Link to the Products's Page"]
expected_unit = first_row["Unit of Measure"]
expected_name = "Replacement Keyed Lock Cores"
expected_desc = "Replacement lock and key set for Bobrick B-3944."

print(f"Testing product: {first_row['Item Number']}")
print(f"Link: {link}")
print(f"Expected Unit: {expected_unit}")
print(f"Expected Name: {expected_name}")
print(f"Expected Description: {expected_desc}")
print("\n" + "="*80)

# Setup Chrome driver
chrome_options = Options()
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_argument('--disable-blink-features=AutomationControlled')
chrome_options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
chrome_options.add_argument('--window-size=1920,1080')

driver = webdriver.Chrome(options=chrome_options)
driver.implicitly_wait(5)

try:
    print("Loading page...")
    driver.get(link)
    
    # Wait for page to load
    time.sleep(3)
    WebDriverWait(driver, 10).until(lambda d: d.execute_script("return document.readyState") == "complete")
    time.sleep(2)
    
    print("\n" + "="*80)
    print("SEARCHING FOR UNIT...")
    print("="*80)
    
    # Get page source
    page_source = driver.page_source
    
    # Search for unit pattern
    print("\n1. Searching for price with unit in page source...")
    price_patterns = [
        r'\$[\d,]+\.?\d{2}\s*/([A-Z]{2,4})\b',
        r'\$[\d,]+\.?\d*\s*/([A-Z]{2,4})\b',
        r'[\d,]+\.?\d{2}\s*/([A-Z]{2,4})\b',
    ]
    
    for pattern in price_patterns:
        matches = list(re.finditer(pattern, page_source, re.IGNORECASE))
        if matches:
            print(f"   Pattern '{pattern}' found {len(matches)} matches:")
            for i, match in enumerate(matches[:5]):  # Show first 5
                unit = match.group(1).upper()
                context = page_source[max(0, match.start()-30):match.end()+30]
                print(f"   Match {i+1}: Unit='{unit}', Context='{context[:80]}...'")
                if unit == expected_unit:
                    print(f"   ✓ FOUND EXPECTED UNIT '{expected_unit}' at position {match.start()}")
    
    # Search in visible elements
    print("\n2. Searching for price with unit in visible elements...")
    price_elements = driver.find_elements(By.XPATH, "//*[contains(text(), '$') and contains(text(), '/')]")
    print(f"   Found {len(price_elements)} elements with $ and /")
    for i, elem in enumerate(price_elements[:10]):
        text = elem.text.strip()
        if text:
            print(f"   Element {i+1}: '{text[:100]}'")
            if expected_unit in text.upper():
                print(f"   ✓ FOUND EXPECTED UNIT '{expected_unit}' in element")
                print(f"   Element tag: {elem.tag_name}")
                print(f"   Element class: {elem.get_attribute('class')}")
                print(f"   Element XPath: {driver.execute_script('return arguments[0].getNodeName() + getXPath(arguments[0])', elem) if 'getXPath' in dir() else 'N/A'}")
    
    print("\n" + "="*80)
    print("SEARCHING FOR PRODUCT NAME...")
    print("="*80)
    
    # Search for product name
    print("\n1. Searching for 'Global Product Type' in page source...")
    if 'Global Product Type' in page_source:
        print("   Found 'Global Product Type' in page source")
        # Try to find the value
        patterns = [
            r'Global Product Type[:\s]+([^\n<]+)',
            r'Global Product Type[:\s]*</[^>]*>([^<]+)',
        ]
        for pattern in patterns:
            match = re.search(pattern, page_source, re.IGNORECASE)
            if match:
                found_name = match.group(1).strip()
                found_name = re.sub(r'<[^>]+>', '', found_name).strip()
                print(f"   Pattern '{pattern}' found: '{found_name}'")
                if expected_name.lower() in found_name.lower() or found_name.lower() in expected_name.lower():
                    print(f"   ✓ MATCHES EXPECTED NAME")
    
    print("\n2. Searching for product name in visible elements...")
    name_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'Global Product Type')]")
    print(f"   Found {len(name_elements)} elements with 'Global Product Type'")
    for i, elem in enumerate(name_elements[:5]):
        print(f"   Element {i+1}: '{elem.text[:100]}'")
        # Try to find following sibling
        try:
            parent = elem.find_element(By.XPATH, "./..")
            print(f"   Parent text: '{parent.text[:200]}'")
        except:
            pass
    
    # Search for the exact name
    if expected_name:
        name_elements = driver.find_elements(By.XPATH, f"//*[contains(text(), '{expected_name}')]")
        print(f"\n   Found {len(name_elements)} elements containing expected name")
        for i, elem in enumerate(name_elements[:3]):
            print(f"   Element {i+1}: Tag={elem.tag_name}, Text='{elem.text[:100]}'")
    
    print("\n" + "="*80)
    print("SEARCHING FOR DESCRIPTION...")
    print("="*80)
    
    # Search for description
    print("\n1. Searching for description in page source...")
    if 'Description' in page_source:
        print("   Found 'Description' in page source")
        # Try to find description content
        desc_patterns = [
            r'Description[:\s]+([^\n<]{20,})',
            r'<[^>]*description[^>]*>([^<]+(?:<[^>]+>[^<]+)*)</[^>]*>',
        ]
        for pattern in desc_patterns:
            matches = list(re.finditer(pattern, page_source, re.IGNORECASE | re.DOTALL))
            if matches:
                print(f"   Pattern '{pattern}' found {len(matches)} matches:")
                for i, match in enumerate(matches[:3]):
                    desc = match.group(1).strip()
                    desc = re.sub(r'<[^>]+>', ' ', desc)
                    desc = ' '.join(desc.split())
                    if len(desc) > 20:
                        print(f"   Match {i+1}: '{desc[:150]}...'")
                        if expected_desc.lower() in desc.lower():
                            print(f"   ✓ MATCHES EXPECTED DESCRIPTION")
    
    print("\n2. Searching for description in visible elements...")
    desc_elements = driver.find_elements(By.XPATH, "//*[contains(@class, 'description') or contains(text(), 'Description')]")
    print(f"   Found {len(desc_elements)} elements with 'description'")
    for i, elem in enumerate(desc_elements[:5]):
        text = elem.text.strip()
        if text and len(text) > 20:
            print(f"   Element {i+1}: '{text[:150]}...'")
            if expected_desc.lower() in text.lower():
                print(f"   ✓ MATCHES EXPECTED DESCRIPTION")
                print(f"   Element tag: {elem.tag_name}")
                print(f"   Element class: {elem.get_attribute('class')}")
    
    # Search for exact description text
    if expected_desc:
        desc_keywords = expected_desc.split()[:3]  # First 3 words
        search_text = ' '.join(desc_keywords)
        desc_elements = driver.find_elements(By.XPATH, f"//*[contains(text(), '{search_text}')]")
        print(f"\n   Found {len(desc_elements)} elements containing '{search_text}'")
        for i, elem in enumerate(desc_elements[:3]):
            print(f"   Element {i+1}: Tag={elem.tag_name}, Text='{elem.text[:150]}...'")
    
    print("\n" + "="*80)
    print("PAGE SOURCE SAMPLE (around price area)...")
    print("="*80)
    # Find price in source and show context
    price_match = re.search(r'\$[\d,]+\.?\d*', page_source)
    if price_match:
        start = max(0, price_match.start() - 100)
        end = min(len(page_source), price_match.end() + 200)
        print(page_source[start:end])
    
    input("\n\nPress Enter to close browser...")
    
finally:
    driver.quit()

