import pandas as pd
import os
import sys
import re
import time
import shutil
import json
from datetime import datetime
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, InvalidSessionIdException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from urllib3.exceptions import ReadTimeoutError, ConnectionError as Urllib3ConnectionError
import socket
import html

# Debug logging helper
DEBUG_LOG_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), ".cursor", "debug.log")
def debug_log(location, message, data=None, hypothesis_id=None):
    try:
        log_entry = {
            "timestamp": int(time.time() * 1000),
            "location": location,
            "message": message,
            "data": data or {},
            "sessionId": "debug-session",
            "runId": "run1",
            "hypothesisId": hypothesis_id
        }
        os.makedirs(os.path.dirname(DEBUG_LOG_PATH), exist_ok=True)
        with open(DEBUG_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(json.dumps(log_entry) + "\n")
        # Also print for immediate visibility
        print(f"[DEBUG] {location}: {message} - {json.dumps(data) if data else ''}")
    except Exception as e:
        print(f"[DEBUG LOG ERROR] {e}")

# Path to Excel file (in parent folder)
excel_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "ScrappedProducts.xlsx")
backup_folder = os.path.join(os.path.dirname(os.path.dirname(__file__)), "Backups")
MAX_BACKUPS = 5  # Keep last 5 backups

# Read the Excel file
print("Reading Excel file...")
try:
    # Check if file exists
    if not os.path.exists(excel_path):
        print(f"ERROR: Excel file not found at: {excel_path}")
        print("Please make sure the file exists and the path is correct.")
        sys.exit(1)
    
    # Check if file is accessible (not locked by another program)
    try:
        # Try to open the file to check if it's locked
        test_file = open(excel_path, 'r+b')
        test_file.close()
    except PermissionError:
        print(f"ERROR: Excel file is locked or in use!")
        print("Please close the Excel file if it's open and try again.")
        sys.exit(1)
    except Exception as e:
        print(f"ERROR: Cannot access Excel file: {e}")
        sys.exit(1)
    
    # Try to read the Excel file
    df = pd.read_excel(excel_path)
    print(f"Successfully loaded Excel file with {len(df)} rows")
    
except pd.errors.EmptyDataError:
    print(f"ERROR: Excel file is empty: {excel_path}")
    sys.exit(1)
except Exception as e:
    error_msg = str(e).lower()
    if 'badzipfile' in error_msg or 'not a zip file' in error_msg:
        print(f"ERROR: Excel file appears to be corrupted or not a valid Excel file!")
        print(f"File path: {excel_path}")
        print("Possible causes:")
        print("  1. File is corrupted - try opening it in Excel and saving again")
        print("  2. File is currently open in Excel - close it and try again")
        print("  3. File is not actually an Excel file - check the file extension")
        print(f"\nError details: {e}")
    else:
        print(f"ERROR: Failed to read Excel file: {e}")
        print(f"File path: {excel_path}")
    sys.exit(1)

# Find the required columns
link_col = None
unit_col = None
product_name_col = None
description_col = None
image_url_col = None

for col in df.columns:
    col_lower = col.lower()
    if 'link' in col_lower and 'product' in col_lower:
        link_col = col
    elif 'unit' in col_lower and 'measure' in col_lower:
        unit_col = col
    elif 'product' in col_lower and 'name' in col_lower:
        product_name_col = col
    elif 'description' in col_lower:
        description_col = col
    elif 'image' in col_lower and 'url' in col_lower:
        image_url_col = col

print(f"Link column: {link_col}")
print(f"Unit column: {unit_col}")
print(f"Product Name column: {product_name_col}")
print(f"Description column: {description_col}")
print(f"Image URL column: {image_url_col}")

# Global driver variable
driver = None

def create_backup():
    """Create a backup of the Excel file and keep only the last MAX_BACKUPS backups"""
    try:
        # Create backup folder if it doesn't exist
        os.makedirs(backup_folder, exist_ok=True)
        
        # Check if Excel file exists
        if not os.path.exists(excel_path):
            print("Warning: Excel file not found, skipping backup...")
            return
        
        # Create backup filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"ScrappedProducts_backup_{timestamp}.xlsx"
        backup_path = os.path.join(backup_folder, backup_filename)
        
        # Copy the Excel file to backup location
        shutil.copy2(excel_path, backup_path)
        print(f"Backup created: {backup_filename}")
        
        # Get all backup files and sort by modification time (newest first)
        backup_files = []
        for file in os.listdir(backup_folder):
            if file.startswith("ScrappedProducts_backup_") and file.endswith(".xlsx"):
                file_path = os.path.join(backup_folder, file)
                backup_files.append((file_path, os.path.getmtime(file_path)))
        
        # Sort by modification time (newest first)
        backup_files.sort(key=lambda x: x[1], reverse=True)
        
        # Keep only the last MAX_BACKUPS backups
        if len(backup_files) > MAX_BACKUPS:
            # Delete older backups
            for file_path, _ in backup_files[MAX_BACKUPS:]:
                try:
                    os.remove(file_path)
                    print(f"Deleted old backup: {os.path.basename(file_path)}")
                except Exception as e:
                    print(f"Error deleting old backup {file_path}: {e}")
        
        # Show current backup count
        remaining_backups = min(len(backup_files), MAX_BACKUPS)
        print(f"Total backups kept: {remaining_backups}/{MAX_BACKUPS}")
        
    except Exception as e:
        print(f"Error creating backup: {e}")

def setup_driver():
    """Setup and return Chrome driver"""
    global driver
    if driver is None:
        print("\nSetting up Chrome driver...")
        chrome_options = Options()
        
        # Set to False to see the browser (useful for debugging)
        HEADLESS_MODE = False
        
        if HEADLESS_MODE:
            chrome_options.add_argument('--headless')  # Run in background
        
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
        chrome_options.add_argument('--window-size=1920,1080')
        
        # Suppress console errors and warnings to speed up execution
        chrome_options.add_argument('--log-level=3')  # Only show fatal errors
        chrome_options.add_argument('--disable-logging')
        chrome_options.add_argument('--disable-gpu-logging')
        chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # Disable background services that cause DEPRECATED_ENDPOINT errors
        chrome_options.add_argument('--disable-background-networking')
        chrome_options.add_argument('--disable-background-timer-throttling')
        chrome_options.add_argument('--disable-backgrounding-occluded-windows')
        chrome_options.add_argument('--disable-breakpad')
        chrome_options.add_argument('--disable-client-side-phishing-detection')
        chrome_options.add_argument('--disable-component-extensions-with-background-page')
        chrome_options.add_argument('--disable-default-apps')
        chrome_options.add_argument('--disable-extensions')
        chrome_options.add_argument('--disable-hang-monitor')
        chrome_options.add_argument('--disable-popup-blocking')
        chrome_options.add_argument('--disable-prompt-on-repost')
        chrome_options.add_argument('--disable-sync')
        chrome_options.add_argument('--disable-translate')
        chrome_options.add_argument('--metrics-recording-only')
        chrome_options.add_argument('--no-first-run')
        chrome_options.add_argument('--safebrowsing-disable-auto-update')
        chrome_options.add_argument('--enable-automation')
        chrome_options.add_argument('--password-store=basic')
        chrome_options.add_argument('--use-mock-keychain')
        
        # Additional options to prevent getting stuck:
        # Disable images to reduce load time and prevent hanging on slow image loads
        chrome_options.add_argument('--blink-settings=imagesEnabled=false')
        chrome_options.add_argument('--disable-plugins')  # Disable plugins
        chrome_options.add_argument('--disable-software-rasterizer')  # Reduce GPU issues
        chrome_options.add_argument('--disable-web-security')  # Sometimes helps with CORS issues
        chrome_options.add_argument('--disable-features=TranslateUI')  # Disable translation UI
        chrome_options.add_argument('--disable-ipc-flooding-protection')  # Prevent IPC issues
        
        # Set preferences to limit resource loading
        prefs = {
            "profile.managed_default_content_settings.images": 2,  # Block images
            "profile.default_content_setting_values.notifications": 2,  # Block notifications
            "profile.default_content_settings.popups": 0,  # Allow popups (might be needed)
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
        driver = webdriver.Chrome(options=chrome_options)
        driver.implicitly_wait(1)  # Reduced to 1 second for faster failure detection
        # Set page load timeout to prevent hanging (15 seconds max - optimized for speed)
        driver.set_page_load_timeout(15)
        # Set script timeout to prevent JS from hanging
        driver.set_script_timeout(15)
    return driver

def recreate_driver():
    """Recreate the Chrome driver when session is lost"""
    global driver
    try:
        if driver is not None:
            try:
                driver.quit()
            except:
                pass  # Ignore errors when closing old driver
    except:
        pass
    
    driver = None
    print("\nRecreating Chrome driver (session was lost)...")
    time.sleep(1)  # Brief pause before recreating (reduced from 2s)
    return setup_driver()

def extract_unit_from_price(price_text):
    """Extract unit from price text like '$1,053.27 /EA'"""
    if not price_text:
        return None
    
    # Common unit abbreviations
    common_units = ['EA', 'BX', 'CS', 'PK', 'CT', 'DZ', 'PR', 'RL', 'FT', 'YD', 'LB', 'OZ', 'GA', 'QT', 'PT', 'FL', 'PC', 'SET', 'PAIR', 'PKG', 'CASE', 'PACK', 'ROLL', 'TUBE', 'BAG', 'BOX', 'CARTON', 'PKT', 'BTL', 'CAN', 'JAR', 'TIN']
    # Common file extensions to exclude
    file_extensions = ['SVG', 'PNG', 'JPG', 'JPEG', 'GIF', 'PDF', 'XML', 'HTML', 'CSS', 'JS', 'JSON']
    
    # Look for pattern like /EA, /BX, /CS, etc. with word boundary
    # Prioritize patterns that look like prices (contain $ or numbers)
    if '$' in price_text or re.search(r'\d', price_text):
        match = re.search(r'/([A-Z]{2,4})\b', price_text.upper())
        if match:
            unit = match.group(1)
            # Prefer common units
            if unit in common_units:
                return unit
            # Also accept 2-3 char units that aren't file extensions
            elif len(unit) >= 2 and len(unit) <= 3 and unit not in file_extensions:
                return unit
    return None

def scrape_product_data(link, expected_unit, retry_count=0):
    """Scrape product data from the webpage - optimized for speed
    
    Args:
        link: URL to scrape
        expected_unit: Expected unit of measure
        retry_count: Internal counter to prevent infinite recursion (max 1 retry)
    """
    try:
        print(f"  Accessing: {link}")
        # Use set_page_load_timeout to prevent hanging (already set in setup_driver, but ensure it's active)
        try:
            driver.set_page_load_timeout(15)  # 15 second timeout for page load (optimized)
            driver.get(link)
        except TimeoutException:
            print(f"    Page load timeout (15s) - stopping page load and skipping")
            try:
                # Stop the page from loading to prevent browser from getting stuck
                driver.execute_script("window.stop();")
            except:
                pass  # Ignore if we can't execute script (browser might be unresponsive)
            try:
                # Navigate to about:blank to reset browser state
                driver.get("about:blank")
            except:
                pass  # Ignore if navigation fails
            time.sleep(0.2)  # Brief pause to let browser recover (reduced from 0.5s)
            return "Timeout error", "Timeout error", "Timeout error", None
        
        # ULTRA-FAST DETECTION: Use JavaScript to check for body (faster than WebDriverWait)
        try:
            # Use JavaScript check for faster detection
            body_exists = driver.execute_script("return document.body !== null;")
            if not body_exists:
                # Fallback to WebDriverWait with shorter timeout
                WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        except:
            print(f"    Product not found (page load timeout)")
            return "Product not found", "Product not found", "Product not found", None
        
        # IMMEDIATE CHECK: Look for key product elements (fastest way to detect if product exists)
        # Cache page_source to avoid multiple fetches
        page_source = driver.page_source
        page_source_lower = page_source.lower()
        
        # Quick check for "not found" indicators first (fastest check)
        not_found_indicators = [
            'product not found', 'item not found', '404', 'page not found', 
            'unavailable', 'error 404', 'item unavailable', 'product unavailable'
        ]
        for indicator in not_found_indicators:
            if indicator in page_source_lower:
                print(f"    Product not found (detected: '{indicator}')")
                return "Product not found", "Product not found", "Product not found", None
        
        # Try to find the unit element with a very short timeout (0.5 second - optimized)
        # This is the fastest way to confirm product exists
        try:
            uom_element = WebDriverWait(driver, 0.5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "span.ess-detail-uom, .ess-detail-uom"))
            )
            # If we found it, product exists - continue to extraction
        except:
            # Unit element not found - do quick checks to confirm product doesn't exist
            # Use JavaScript for faster checking (no need to wait for full page load)
            try:
                # Fast JavaScript check for unit element
                uom_exists = driver.execute_script(
                    "return document.querySelector('span.ess-detail-uom, .ess-detail-uom') !== null;"
                )
                if uom_exists:
                    uom_element = driver.find_element(By.CSS_SELECTOR, "span.ess-detail-uom, .ess-detail-uom")
                else:
                    raise NoSuchElementException("Unit element not found")
            except:
                # Check for key product elements - if none found, product doesn't exist
                product_indicators = [
                    'ess-detail', 'ess-product', 'product-detail', 'item-detail',
                    'product-name', 'product-type', 'global product type'
                ]
                has_product_indicator = any(indicator in page_source_lower for indicator in product_indicators)
                
                # Also check if there's a price with unit pattern (real products have this)
                has_price_with_unit = bool(re.search(r'\$[\d,]+\.?\d*\s*/([A-Z]{2,4})\b', page_source, re.IGNORECASE))
                
                if not has_product_indicator and not has_price_with_unit:
                    print(f"    Product not found (no product elements detected)")
                    return "Product not found", "Product not found", "Product not found", None
                
                # If we have indicators but no unit element, might be loading - try one more quick check (no wait)
                try:
                    uom_element = driver.find_element(By.CSS_SELECTOR, "span.ess-detail-uom, .ess-detail-uom")
                except:
                    # Still can't find unit element - likely product not available
                    print(f"    Product not found (unit element not found)")
                    return "Product not found", "Product not found", "Product not found", None
        
        # Extract unit - we already found the element in the check above, so use it
        website_unit = None
        
        # Common unit abbreviations (EA, BX, CS, PK, CT, DZ, PR, etc.)
        common_units = ['EA', 'BX', 'CS', 'PK', 'CT', 'DZ', 'PR', 'RL', 'FT', 'YD', 'LB', 'OZ', 'GA', 'QT', 'PT', 'FL', 'PC', 'SET', 'PAIR', 'PKG', 'CASE', 'PACK', 'ROLL', 'TUBE', 'BAG', 'BOX', 'CARTON', 'PKT', 'BTL', 'CAN', 'JAR', 'TIN']
        # Common file extensions to exclude
        file_extensions = ['SVG', 'PNG', 'JPG', 'JPEG', 'GIF', 'PDF', 'XML', 'HTML', 'CSS', 'JS', 'JSON']
        
        # Pre-compile regex pattern for faster matching
        uom_pattern = re.compile(r'class="ess-detail-uom"[^>]*>/([A-Z]{2,4})\b', re.IGNORECASE)
        
        try:
            # We already found the unit element in the check above, extract from it
            try:
                uom_element = driver.find_element(By.CSS_SELECTOR, "span.ess-detail-uom, .ess-detail-uom")
                uom_text = uom_element.text.strip()
                # Extract unit from text like "/EA" or "EA"
                if uom_text.startswith('/'):
                    website_unit = uom_text[1:].strip().upper()
                else:
                    website_unit = uom_text.upper()
                if website_unit and len(website_unit) >= 2 and len(website_unit) <= 4:
                    print(f"    Found unit: {website_unit}")
            except:
                # Fallback: Try CSS selector with class contains (faster than XPath)
                try:
                    uom_elements = driver.find_elements(By.CSS_SELECTOR, "[class*='ess-detail-uom'], [class*='ess-product-uom']")
                    for elem in uom_elements:
                        text = elem.text.strip()
                        if text.startswith('/'):
                            unit = text[1:].strip().upper()
                            if unit and len(unit) >= 2 and len(unit) <= 4:
                                website_unit = unit
                                print(f"    Found unit (CSS): {website_unit}")
                                break
                except:
                    pass
            
            # If still not found, try page source regex (quick check using cached page_source)
            if not website_unit:
                match = uom_pattern.search(page_source)
                if match:
                    potential_unit = match.group(1).upper()
                    if potential_unit in common_units or (len(potential_unit) >= 2 and len(potential_unit) <= 4 and potential_unit not in file_extensions):
                        website_unit = potential_unit
                        print(f"    Found unit (regex): {website_unit}")
            
            # If still no unit, product likely not available
            if not website_unit:
                print(f"    Product not found (unit not extractable)")
                return "Product not found", "Product not found", "Product not found", None
                
        except Exception as e:
            print(f"    Error finding unit: {e}")
            return "Product not found", "Product not found", "Product not found", None
        
        print(f"    Website unit: {website_unit}, Expected unit: {expected_unit}")
        
        # Check if unit matches
        if website_unit.upper() != str(expected_unit).upper().strip():
            print(f"    ⚠️  Unit mismatch! Website: {website_unit}, Expected: {expected_unit}")
            return "Unit not matched", "Unit not matched", "Unit not matched", None
        
        # Unit matches, proceed with scraping
        print(f"    ✅ Unit matched! Scraping data...")
        
        # 1. Scrape Image URL - OPTIMIZED for speed
        image_url = None
        try:
            # Pre-compile regex patterns for faster matching
            img_tag_pattern_oppictures = re.compile(r'<img[^>]+src=["\']([^"\']*oppictures[^"\']+)["\']', re.IGNORECASE)
            img_tag_pattern_general = re.compile(r'<img[^>]+src=["\']([^"\']+)["\']', re.IGNORECASE)
            
            # Helper function to check if URL is a product image (prioritize these)
            def is_product_image_url(url):
                """Check if URL is a product image (not a tag/rebate image)"""
                if not url:
                    return False
                url_lower = url.lower()
                # Exclude tag/rebate images
                if '/tags/' in url_lower or 'tags/' in url_lower or 'tagoutlined' in url_lower:
                    return False
                # Prioritize actual product images from Master_Variants
                if 'master_variants' in url_lower or 'variant_' in url_lower:
                    return True
                # Also accept other oppictures images that aren't tags
                if 'oppictures.com' in url_lower and '/tags/' not in url_lower:
                    return True
                return False
            
            # METHOD 1: Use JavaScript to quickly find product images (faster than Selenium)
            try:
                # Use JavaScript to find images with oppictures in src (much faster)
                img_srcs = driver.execute_script("""
                    var imgs = document.querySelectorAll('img[src*="oppictures"]');
                    var srcs = [];
                    for (var i = 0; i < imgs.length; i++) {
                        var src = imgs[i].src;
                        if (src && src.indexOf('/tags/') === -1 && src.indexOf('tagoutlined') === -1) {
                            if (src.indexOf('master_variants') !== -1 || src.indexOf('variant_') !== -1) {
                                srcs.unshift(src); // Prioritize product images
                            } else {
                                srcs.push(src);
                            }
                        }
                    }
                    return srcs;
                """)
                
                if img_srcs and len(img_srcs) > 0:
                    image_url = img_srcs[0]
                    print(f"    Found product image (JS): {image_url[:80]}...")
            except Exception as e:
                pass  # Fallback to other methods
            
            # METHOD 2: Fallback to CSS selector (faster than XPath)
            if not image_url:
                try:
                    img_elements = driver.find_elements(By.CSS_SELECTOR, "img[src*='oppictures']")
                    for img in img_elements:
                        src = img.get_attribute('src')
                        if src and is_product_image_url(src):
                            if src.startswith('//'):
                                src = 'https:' + src
                            image_url = src
                            print(f"    Found product image (CSS): {image_url[:80]}...")
                            break
                except Exception as e:
                    pass
            
            # METHOD 3: Fallback to page source regex (using cached page_source)
            if not image_url:
                # First, prioritize product images from Master_Variants
                matches = img_tag_pattern_oppictures.finditer(page_source)
                product_image_candidates = []
                other_image_candidates = []
                
                for match in matches:
                    src = match.group(1)
                    if src.startswith('//'):
                        src = 'https:' + src
                    # Skip tag images
                    if '/tags/' in src.lower() or 'tags/' in src.lower() or 'tagoutlined' in src.lower():
                        continue
                    if is_product_image_url(src):
                        product_image_candidates.append(src)
                    elif 'oppictures.com' in src.lower():
                        other_image_candidates.append(src)
                
                # Use product images first
                if product_image_candidates:
                    image_url = product_image_candidates[0]
                    print(f"    Found product image from page source (Master_Variants): {image_url[:80]}...")
                elif other_image_candidates:
                    # Filter out any tag images that might have slipped through
                    filtered_candidates = [c for c in other_image_candidates if '/tags/' not in c.lower() and 'tagoutlined' not in c.lower()]
                    if filtered_candidates:
                        image_url = filtered_candidates[0]
                        print(f"    Found image from page source (oppictures): {image_url[:80]}...")
                                
        except Exception as e:
            print(f"    Error finding image: {e}")
        
        # 2. Scrape Product Name (Global Product Type from Product Details section)
        # OPTIMIZED: Use JavaScript for faster extraction
        product_name = None
        try:
            # METHOD 1: Use JavaScript to quickly find product name (faster than XPath)
            try:
                product_name = driver.execute_script("""
                    var tds = document.querySelectorAll('td');
                    for (var i = 0; i < tds.length; i++) {
                        if (tds[i].textContent && tds[i].textContent.trim().indexOf('Global Product Type') !== -1) {
                            var nextTd = tds[i].nextElementSibling;
                            if (nextTd && nextTd.textContent) {
                                var name = nextTd.textContent.trim();
                                if (name && name !== 'Global Product Type' && name.length > 5) {
                                    return name;
                                }
                            }
                        }
                    }
                    return null;
                """)
                if product_name:
                    print(f"    Found product name (JS): {product_name}")
            except:
                pass
            
            # METHOD 2: Fallback to XPath (CSS :contains() is not standard)
            if not product_name:
                try:
                    # Find the td containing "Global Product Type" and get the following sibling td
                    name_element = driver.find_element(By.XPATH, "//td[contains(text(), 'Global Product Type')]/following-sibling::td[1]")
                    product_name = name_element.text.strip()
                    if product_name:
                        print(f"    Found product name from td: {product_name}")
                except:
                    # Alternative: Find all tds and search
                    try:
                        name_elements = driver.find_elements(By.CSS_SELECTOR, "td")
                        for i, elem in enumerate(name_elements):
                            text = elem.text.strip()
                            if 'Global Product Type' in text and i + 1 < len(name_elements):
                                next_text = name_elements[i + 1].text.strip()
                                if next_text and next_text != 'Global Product Type' and len(next_text) > 5:
                                    product_name = next_text
                                    print(f"    Found product name from alternative td: {product_name}")
                                    break
                    except:
                        pass
            
            # METHOD 3: Fallback to page source pattern (using cached page_source)
            if not product_name:
                pattern = re.compile(r'Global Product Type[:\s]+([^\n<]+)', re.IGNORECASE)
                match = pattern.search(page_source)
                if match:
                    product_name = match.group(1).strip()
                    product_name = re.sub(r'<[^>]+>', '', product_name).strip()
                    if product_name:
                        print(f"    Found product name from page source: {product_name}")
                    
        except Exception as e:
            print(f"    Error finding product name: {e}")
        
        # 3. Scrape Description - OPTIMIZED for speed
        # Extract only the actual product description, excluding warnings, recommendations, pricing, and UI elements
        description = None
        
        # Pre-compile regex patterns for faster matching
        html_tag_pattern = re.compile(r'<[^>]+>')
        desc_pattern = re.compile(r'description\s*:?\s*', re.IGNORECASE)
        stop_markers_pattern = re.compile(r'(Product Details|ADD TO LIST|People Who Bought|Also Consider|List price)', re.IGNORECASE)
        price_pattern = re.compile(r'\$\d+[.,]\d+\s*/[A-Z]{2,4}', re.IGNORECASE)
        
        def clean_description_text(text):
            """Clean description text by removing HTML fragments, section markers, and unwanted content"""
            if not text:
                return None
            
            # Remove any remaining HTML tags and fragments
            text = html_tag_pattern.sub(' ', text)
            text = html.unescape(text)
            
            # Remove section markers and UI elements at the end
            text = stop_markers_pattern.sub('', text)
            
            # Remove the "Description :" or "Description:" prefix
            text = desc_pattern.sub('', text, count=1)
            
            # Remove item numbers at the start
            text = re.sub(r'^[A-Z0-9]{6,15}\s+', '', text)
            
            # Remove price patterns
            text = price_pattern.sub('', text)
            
            # Normalize whitespace
            text = ' '.join(text.split())
            text = text.strip()
            
            # Remove trailing punctuation
            text = re.sub(r'\*+\s*$', '', text)
            text = text.strip()
            
            return text if text else None
        
        try:
            # METHOD 1: Use JavaScript to quickly find description (faster than DOM traversal)
            try:
                description = driver.execute_script("""
                    var elements = document.querySelectorAll('*');
                    for (var i = 0; i < elements.length; i++) {
                        var text = elements[i].textContent || '';
                        if (text.indexOf('Description') !== -1 && (text.indexOf(':') !== -1 || text.indexOf(' :') !== -1)) {
                            // Try to get text after "Description :"
                            var match = text.match(/Description\\s*:?\\s*(.+?)(?:Product Details|ADD TO LIST|People Who|List price)/i);
                            if (match && match[1]) {
                                var desc = match[1].trim();
                                if (desc.length >= 20 && desc.length <= 10000) {
                                    return desc;
                                }
                            }
                        }
                    }
                    return null;
                """)
                if description:
                    description = clean_description_text(description)
                    if description and 20 <= len(description) <= 10000:
                        print(f"    Found description (JS): {len(description)} chars")
                    else:
                        description = None
            except:
                pass
            
            # METHOD 2: Fallback to page source pattern (using cached page_source)
            if not description:
                desc_heading_patterns = ['Description :', 'Description:', 'Description']
                desc_index = -1
                for pattern in desc_heading_patterns:
                    desc_index = page_source.find(pattern)
                    if desc_index != -1:
                        break
                
                if desc_index != -1:
                    # Get text chunk after the description heading
                    text_chunk = page_source[desc_index:desc_index + 10000]
                    
                    # Remove HTML tags and decode HTML entities
                    text_only = html_tag_pattern.sub(' ', text_chunk)
                    text_only = html.unescape(text_only)
                    text_only = ' '.join(text_only.split())
                    
                    # Remove the "Description :" prefix
                    text_only = desc_pattern.sub('', text_only, count=1)
                    text_only = text_only.strip()
                    
                    # Find stop markers and truncate if in last 30% of text
                    text_length = len(text_only)
                    stop_match = stop_markers_pattern.search(text_only)
                    if stop_match and stop_match.start() > text_length * 0.7:
                        text_only = text_only[:stop_match.start()].strip()
                    
                    # Clean the description text
                    text_only = clean_description_text(text_only)
                    
                    # Final validation
                    if text_only and 20 <= len(text_only) <= 10000:
                        description = text_only
                        print(f"    Found description (page source): {len(description)} chars")
        
        except Exception as e:
            pass  # Silently continue
        return product_name, description, image_url, website_unit
        
    except TimeoutException:
        print(f"    Page load timeout (30s) - stopping page load")
        try:
            # Stop the page from loading to prevent browser from getting stuck
            driver.execute_script("window.stop();")
        except:
            pass  # Ignore if we can't execute script
        try:
            # Navigate to about:blank to reset browser state
            driver.get("about:blank")
        except:
            pass  # Ignore if navigation fails
        time.sleep(0.2)  # Brief pause to let browser recover (optimized)
        return "Timeout error", "Timeout error", "Timeout error", None
    except (ReadTimeoutError, Urllib3ConnectionError, socket.timeout, ConnectionResetError) as e:
        error_msg = str(e).lower()
        if 'read timeout' in error_msg or 'connection' in error_msg or 'timeout' in error_msg:
            print(f"    Connection timeout/error - marking as timeout")
            return "Timeout error", "Timeout error", "Timeout error", None
        else:
            print(f"    Connection error: {e}")
            return "Timeout error", "Timeout error", "Timeout error", None
    except (InvalidSessionIdException, WebDriverException) as e:
        error_msg = str(e).lower()
        if 'invalid session id' in error_msg or 'session id' in error_msg:
            if retry_count < 1:  # Only retry once
                print(f"    Browser session lost. Recreating driver and retrying...")
                try:
                    recreate_driver()
                    # Retry the entire scraping operation
                    print(f"    Retrying: {link}")
                    return scrape_product_data(link, expected_unit, retry_count + 1)  # Recursive retry
                except Exception as retry_error:
                    print(f"    Failed to recreate driver or retry failed: {retry_error}")
                    return None, None, None, None
            else:
                print(f"    Browser session lost (max retries reached)")
                return None, None, None, None
        else:
            print(f"    WebDriver error: {e}")
            return None, None, None, None
    except Exception as e:
        error_msg = str(e).lower()
        # Check for invalid session id in error message (sometimes it's wrapped in a generic exception)
        if 'invalid session id' in error_msg or 'session id' in error_msg:
            if retry_count < 1:  # Only retry once
                print(f"    Browser session lost. Recreating driver and retrying...")
                try:
                    recreate_driver()
                    # Retry the entire scraping operation
                    print(f"    Retrying: {link}")
                    return scrape_product_data(link, expected_unit, retry_count + 1)  # Recursive retry
                except Exception as retry_error:
                    print(f"    Failed to recreate driver or retry failed: {retry_error}")
                    return None, None, None, None
            else:
                print(f"    Browser session lost (max retries reached)")
                return None, None, None, None
        # Check for timeout/connection errors in generic exceptions
        elif 'timeout' in error_msg or 'read timeout' in error_msg or 'connection' in error_msg:
            print(f"    Timeout/connection error detected")
            return "Timeout error", "Timeout error", "Timeout error", None
        else:
            print(f"    Error scraping: {e}")
            return None, None, None, None

def process_products(start_idx=0, end_idx=None, test_mode=False):
    """Process products from start_idx to end_idx (inclusive)
    
    Args:
        start_idx: Starting row index (0-based)
        end_idx: Ending row index (0-based, None for last row)
        test_mode: If True, indicates test mode (currently no special behavior)
    """
    global df
    
    # Create backup before starting
    print("\nCreating backup before starting...")
    create_backup()
    
    # Setup driver if not already done
    setup_driver()
    
    if end_idx is None:
        end_idx = len(df) - 1
    
    # Ensure valid range
    start_idx = max(0, start_idx)
    end_idx = min(len(df) - 1, end_idx)
    
    if start_idx > end_idx:
        print("Error: Start index must be less than or equal to end index")
        return
    
    # Convert relevant columns to string type to avoid FutureWarning when writing string values
    if product_name_col in df.columns:
        df[product_name_col] = df[product_name_col].astype(str).replace('nan', '')
    if description_col in df.columns:
        df[description_col] = df[description_col].astype(str).replace('nan', '')
    if image_url_col in df.columns:
        df[image_url_col] = df[image_url_col].astype(str).replace('nan', '')
    
    total_to_process = end_idx - start_idx + 1
    print(f"\nProcessing products from row {start_idx + 1} to {end_idx + 1} ({total_to_process} products)...\n")
    
    processed_count = 0
    skipped_count = 0
    error_count = 0
    
    # Function to safely save progress
    def save_progress_safely():
        """Safely save the Excel file"""
        try:
            print(f"\nSaving progress before exit...")
            df.to_excel(excel_path, index=False)
            print(f"Progress saved successfully! (Processed: {processed_count}, Skipped: {skipped_count}, Errors: {error_count})")
            return True
        except Exception as e:
            print(f"ERROR: Failed to save progress: {e}")
            # Try to save to a backup location
            try:
                backup_path = excel_path.replace('.xlsx', '_emergency_backup.xlsx')
                df.to_excel(backup_path, index=False)
                print(f"Emergency backup saved to: {backup_path}")
                return True
            except Exception as backup_error:
                print(f"ERROR: Failed to create emergency backup: {backup_error}")
                return False
    
    # Find Item Number column for display
    item_number_col = None
    for col in df.columns:
        col_lower = col.lower()
        if col_lower == 'item number':
            item_number_col = col
            break
        elif 'item' in col_lower and 'number' in col_lower:
            if 'stock' not in col_lower and 'butted' not in col_lower:
                item_number_col = col
                break
    
    try:
        for idx in range(start_idx, end_idx + 1):
            row = df.iloc[idx]
            
            # Get item number for display
            item_number = str(row[item_number_col]).strip() if item_number_col and pd.notna(row[item_number_col]) else f"Row {idx + 1}"
            
            print(f"\n{'='*70}")
            print(f"📦 Product: {item_number} | Row {idx + 1}/{len(df)} | Progress: {processed_count + 1}/{total_to_process}")
            print(f"{'='*70}")
            
            # Get link
            link = row[link_col]
            if pd.isna(link) or str(link).strip() == '':
                print("  ❌ No link found, skipping...")
                skipped_count += 1
                continue
            
            # Get expected unit
            expected_unit = row[unit_col]
            if pd.isna(expected_unit):
                print("  ❌ No unit of measure found, skipping...")
                skipped_count += 1
                continue
            
            # Check current status of the three columns
            current_product_name = str(row[product_name_col]).strip() if pd.notna(row[product_name_col]) else ''
            current_description = str(row[description_col]).strip() if pd.notna(row[description_col]) else ''
            current_image_url = str(row[image_url_col]).strip() if pd.notna(row[image_url_col]) else ''
            
            # Helper function to get status emoji
            def get_status_emoji(value, error_messages):
                if not value or value == '':
                    return '⚪', 'Empty'
                elif value in error_messages:
                    if value == 'Product not found':
                        return '❌', 'Product not found'
                    elif value == 'Unit not matched':
                        return '⚠️', 'Unit not matched'
                    elif value == 'Timeout error':
                        return '⏱️', 'Timeout error'
                    else:
                        return '⚠️', value
                else:
                    return '✅', 'Found'
            
            error_messages = ['Unit not matched', 'Product not found', 'Timeout error', '']
            
            # Display current status
            pn_emoji, pn_status = get_status_emoji(current_product_name, error_messages)
            desc_emoji, desc_status = get_status_emoji(current_description, error_messages)
            img_emoji, img_status = get_status_emoji(current_image_url, error_messages)
            
            print(f"\n📊 Current Status:")
            print(f"   {pn_emoji} Product Name: {pn_status}")
            print(f"   {desc_emoji} Description: {desc_status}")
            print(f"   {img_emoji} Image URL: {img_status}")
            
            # Skip if any column has "Product not found" (product wasn't found on website)
            if (current_product_name == 'Product not found' or 
                current_description == 'Product not found' or 
                current_image_url == 'Product not found'):
                print(f"\n  ⏭️  Already marked as 'Product not found', skipping...")
                skipped_count += 1
                continue
            
            # Count how many columns have valid data (not empty, not error messages)
            filled_count = 0
            if current_product_name and current_product_name not in error_messages:
                filled_count += 1
            if current_description and current_description not in error_messages:
                filled_count += 1
            if current_image_url and current_image_url not in error_messages:
                filled_count += 1
            
            # Skip if all 3 columns are already filled
            if filled_count == 3:
                print(f"\n  ✅ All columns already filled, skipping...")
                skipped_count += 1
                continue
            
            # If 1 or 2 columns are filled, we'll re-scrape to fill the empty ones
            if filled_count > 0:
                print(f"\n  🔄 Partial data found ({filled_count}/3 columns filled), re-scraping to fill empty columns...")
            else:
                print(f"\n  🆕 New product, scraping all columns...")
            
            # Scrape data (with retry on session loss, but not for timeout errors)
            print(f"\n  🌐 Accessing: {link}")
            print(f"  🔍 Expected Unit: {expected_unit}")
            
            max_retries = 1  # Reduced retries to avoid getting stuck
            retry_count = 0
            product_name, description, image_url, website_unit = None, None, None, None
            
            while retry_count <= max_retries:
                product_name, description, image_url, website_unit = scrape_product_data(str(link).strip(), expected_unit)
                
                # If we got results (even if error like "Timeout error"), break immediately
                if product_name is not None:
                    break
                
                # If we got None, None, None, None and it might be a session issue, retry once
                retry_count += 1
                if retry_count <= max_retries:
                    print(f"    🔄 Retrying ({retry_count}/{max_retries})...")
                    time.sleep(1)  # Pause before retry (reduced from 2s)
            
            # Display scraping results
            print(f"\n📥 Scraping Results:")
            
            # Update Excel row - only fill empty columns, don't overwrite existing data
            if product_name == "Unit not matched":
                print(f"   ⚠️  Unit not matched! Website unit doesn't match expected unit.")
                # Only update if column is empty or has error message
                updated_pn = not current_product_name or current_product_name in ['Unit not matched', 'Timeout error', '']
                updated_desc = not current_description or current_description in ['Unit not matched', 'Timeout error', '']
                updated_img = not current_image_url or current_image_url in ['Unit not matched', 'Timeout error', '']
                
                if updated_pn:
                    df.at[idx, product_name_col] = "Unit not matched"
                if updated_desc:
                    df.at[idx, description_col] = "Unit not matched"
                if updated_img:
                    df.at[idx, image_url_col] = "Unit not matched"
                
                print(f"   📝 Updated: Product Name: {'✅' if updated_pn else '⏭️'}, Description: {'✅' if updated_desc else '⏭️'}, Image URL: {'✅' if updated_img else '⏭️'}")
                processed_count += 1
            elif product_name == "Product not found":
                print(f"   ❌ Product not found on website!")
                # Only update if column is empty or has error message
                updated_pn = not current_product_name or current_product_name in ['Product not found', 'Timeout error', '']
                updated_desc = not current_description or current_description in ['Product not found', 'Timeout error', '']
                updated_img = not current_image_url or current_image_url in ['Product not found', 'Timeout error', '']
                
                if updated_pn:
                    df.at[idx, product_name_col] = "Product not found"
                if updated_desc:
                    df.at[idx, description_col] = "Product not found"
                if updated_img:
                    df.at[idx, image_url_col] = "Product not found"
                
                print(f"   📝 Updated: Product Name: {'✅' if updated_pn else '⏭️'}, Description: {'✅' if updated_desc else '⏭️'}, Image URL: {'✅' if updated_img else '⏭️'}")
                processed_count += 1
            elif product_name == "Timeout error":
                print(f"   ⏱️  Timeout error occurred!")
                # Only update if column is empty (don't overwrite valid data with timeout error)
                updated_pn = not current_product_name or current_product_name in ['Timeout error', '']
                updated_desc = not current_description or current_description in ['Timeout error', '']
                updated_img = not current_image_url or current_image_url in ['Timeout error', '']
                
                if updated_pn:
                    df.at[idx, product_name_col] = "Timeout error"
                if updated_desc:
                    df.at[idx, description_col] = "Timeout error"
                if updated_img:
                    df.at[idx, image_url_col] = "Timeout error"
                
                print(f"   📝 Updated: Product Name: {'✅' if updated_pn else '⏭️'}, Description: {'✅' if updated_desc else '⏭️'}, Image URL: {'✅' if updated_img else '⏭️'}")
                processed_count += 1
            elif product_name:
                # Display what was found
                pn_found = '✅' if product_name else '❌'
                desc_found = '✅' if description else '❌'
                img_found = '✅' if image_url else '❌'
                
                print(f"   {pn_found} Product Name: {'Found' if product_name else 'Not found'}")
                if product_name:
                    print(f"      └─ {product_name[:60]}{'...' if len(product_name) > 60 else ''}")
                
                print(f"   {desc_found} Description: {'Found' if description else 'Not found'}")
                if description:
                    print(f"      └─ {description[:60]}{'...' if len(description) > 60 else ''} ({len(description)} chars)")
                
                print(f"   {img_found} Image URL: {'Found' if image_url else 'Not found'}")
                if image_url:
                    print(f"      └─ {image_url[:60]}{'...' if len(image_url) > 60 else ''}")
                
                # Only fill empty columns, preserve existing valid data
                updated_pn = not current_product_name or current_product_name in ['Unit not matched', 'Timeout error', '']
                updated_desc = description and (not current_description or current_description in ['Unit not matched', 'Timeout error', ''])
                updated_img = image_url and (not current_image_url or current_image_url in ['Unit not matched', 'Timeout error', ''])
                
                # Update Product Name if empty or has error message
                if updated_pn:
                    df.at[idx, product_name_col] = product_name
                # Update Description if empty or has error message
                if updated_desc:
                    df.at[idx, description_col] = description
                # Update Image URL if empty or has error message
                if updated_img:
                    df.at[idx, image_url_col] = image_url
                
                print(f"\n   📝 Excel Update:")
                print(f"      Product Name: {'✅ Updated' if updated_pn else '⏭️  Preserved (already has data)'}")
                print(f"      Description: {'✅ Updated' if updated_desc else '⏭️  Preserved (already has data)' if current_description and current_description not in error_messages else '❌ Not found'}")
                print(f"      Image URL: {'✅ Updated' if updated_img else '⏭️  Preserved (already has data)' if current_image_url and current_image_url not in error_messages else '❌ Not found'}")
                
                processed_count += 1
            else:
                print(f"   ❌ Error: No data returned from scraper")
                error_count += 1
            
            # Save progress (every 10 products)
            if processed_count % 10 == 0 and processed_count > 0:
                print(f"\nSaving progress...")
                df.to_excel(excel_path, index=False)
                print(f"Progress saved! (Processed: {processed_count}, Skipped: {skipped_count}, Errors: {error_count})\n")
            
            # Small delay to avoid overwhelming the server (optimized)
            time.sleep(0.2)
            
            # Check if browser is still responsive (quick check to prevent getting stuck)
            try:
                driver.current_url  # Simple check to see if browser is responsive
            except (InvalidSessionIdException, WebDriverException, Exception) as e:
                # Browser became unresponsive, recreate driver
                error_msg = str(e).lower()
                if 'timeout' in error_msg or 'session' in error_msg or 'connection' in error_msg:
                    print(f"    Browser unresponsive, recreating driver...")
                    try:
                        recreate_driver()
                    except:
                        print(f"    Failed to recreate driver, will try on next product")
        
        # Final save (normal completion)
        print(f"\nSaving final results...")
        df.to_excel(excel_path, index=False)
        
        # Create backup after completion
        print("Creating backup after completion...")
        create_backup()
        
        print(f"\nDone! Processed: {processed_count}, Skipped: {skipped_count}, Errors: {error_count}")
    
    except KeyboardInterrupt:
        # User pressed Ctrl+C - save progress before exiting
        print("\n\n" + "="*60)
        print("INTERRUPTED BY USER (Ctrl+C)")
        print("="*60)
        print(f"\nSaving progress before exit...")
        save_progress_safely()
        print(f"\nProgress saved! You can resume from where you left off.")
        print(f"Processed: {processed_count}, Skipped: {skipped_count}, Errors: {error_count}")
        print("\nExiting gracefully...")
        raise  # Re-raise to exit the function

def display_menu():
    """Display the main menu"""
    print("\n" + "="*60)
    print("           PRODUCT SCRAPER - MENU")
    print("="*60)
    print("\n1. Scrape ALL products")
    print("2. Scrape first 10 products (TEST)")
    print("3. Scrape products in a range (e.g., 100-2000)")
    print("4. Scrape from a specific product to end")
    print("5. Scrape a single product by Item Number")
    print("6. Exit")
    print("\n" + "="*60)

def get_user_choice():
    """Get user's menu choice"""
    while True:
        try:
            choice = input("\nEnter your choice (1-6): ").strip()
            if choice in ['1', '2', '3', '4', '5', '6']:
                return choice
            else:
                print("Invalid choice. Please enter a number between 1 and 6.")
        except KeyboardInterrupt:
            print("\n\nExiting...")
            return '6'

def get_range_input(total_rows):
    """Get range input from user - asks for start first, then end"""
    # Get starting row number
    while True:
        try:
            start_input = input("Enter starting row number (1-based): ").strip()
            if not start_input:
                print("Error: Starting row number cannot be empty")
                continue
            
            start = int(start_input) - 1  # Convert to 0-based index
            
            if start < 0:
                print("Error: Starting row number must be >= 1")
                continue
            
            if start >= total_rows:
                print(f"Error: Starting row number ({start + 1}) exceeds total rows ({total_rows})")
                continue
            
            break
        except ValueError:
            print("Error: Please enter a valid number")
        except KeyboardInterrupt:
            print("\n\nCancelled.")
            return None, None
    
    # Get ending row number
    while True:
        try:
            end_input = input("Enter ending row number (1-based): ").strip()
            if not end_input:
                print("Error: Ending row number cannot be empty")
                continue
            
            end = int(end_input) - 1  # Convert to 0-based index
            
            if end < 0:
                print("Error: Ending row number must be >= 1")
                continue
            
            if end < start:
                print(f"Error: Ending row number ({end + 1}) must be >= starting row number ({start + 1})")
                continue
            
            if end >= total_rows:
                print(f"Warning: Ending row number ({end + 1}) exceeds total rows ({total_rows}). Using {total_rows} as end.")
                end = total_rows - 1
            
            return start, end
        except ValueError:
            print("Error: Please enter a valid number")
        except KeyboardInterrupt:
            print("\n\nCancelled.")
            return None, None

def get_start_index():
    """Get start index from user"""
    while True:
        try:
            start_input = input("Enter starting row number (1-based): ").strip()
            start = int(start_input) - 1  # Convert to 0-based index
            if start >= 0:
                return start
            else:
                print("Error: Start index must be >= 1")
        except ValueError:
            print("Error: Please enter a valid number")
        except KeyboardInterrupt:
            print("\n\nCancelled.")
            return None

def get_item_number():
    """Get Item Number from user and find its row index"""
    global df
    
    item_number_col = None
    
    # Find Item Number column
    for col in df.columns:
        col_lower = col.lower()
        if col_lower == 'item number':
            item_number_col = col
            break
        elif 'item' in col_lower and 'number' in col_lower:
            if 'stock' not in col_lower and 'butted' not in col_lower:
                item_number_col = col
                break
    
    if item_number_col is None:
        print("Error: Could not find 'Item Number' column")
        return None
    
    while True:
        try:
            item_input = input("Enter Item Number: ").strip()
            if not item_input:
                print("Error: Item Number cannot be empty")
                continue
            
            # Search for the item number in the dataframe
            matching_rows = df[df[item_number_col].astype(str).str.strip() == item_input]
            
            if len(matching_rows) == 0:
                print(f"Error: Item Number '{item_input}' not found in Excel file")
                retry = input("Do you want to try again? (y/n): ").strip().lower()
                if retry != 'y':
                    return None
            elif len(matching_rows) > 1:
                print(f"Warning: Found {len(matching_rows)} rows with Item Number '{item_input}'")
                print("Using the first match...")
                row_idx = matching_rows.index[0]
                return row_idx
            else:
                row_idx = matching_rows.index[0]
                print(f"Found Item Number '{item_input}' at row {row_idx + 1}")
                return row_idx
                
        except KeyboardInterrupt:
            print("\n\nCancelled.")
            return None
        except Exception as e:
            print(f"Error: {e}")
            return None

# Main menu loop
total_rows = len(df)
print(f"\nTotal products in Excel: {total_rows}")

while True:
    display_menu()
    choice = get_user_choice()
    
    if choice == '1':
        # Scrape all products
        confirm = input(f"\nThis will scrape all {total_rows} products. Continue? (y/n): ").strip().lower()
        if confirm == 'y':
            process_products(0, total_rows - 1)
        else:
            print("Cancelled.")
    
    elif choice == '2':
        # Scrape first 10 products (test)
        print("\nScraping first 10 products for testing...")
        process_products(0, min(9, total_rows - 1), test_mode=True)
    
    elif choice == '3':
        # Scrape products in a range
        start, end = get_range_input(total_rows)
        if start is not None and end is not None:
            process_products(start, end)
    
    elif choice == '4':
        # Scrape from specific product to end
        start = get_start_index()
        if start is not None:
            if start >= total_rows:
                print(f"Error: Start index ({start + 1}) exceeds total rows ({total_rows})")
            else:
                process_products(start, total_rows - 1)
    
    elif choice == '5':
        # Scrape a single product by Item Number
        row_idx = get_item_number()
        if row_idx is not None:
            print(f"\nScraping product at row {row_idx + 1}...")
            process_products(row_idx, row_idx, test_mode=True)
    
    elif choice == '6':
        # Exit
        print("\nExiting...")
        break
    
    # Ask if user wants to continue
    if choice != '6':
        continue_choice = input("\nDo you want to perform another operation? (y/n): ").strip().lower()
        if continue_choice != 'y':
            break

# Close driver if it was opened
if driver is not None:
    print("\nClosing browser...")
    driver.quit()
print("Goodbye!")

