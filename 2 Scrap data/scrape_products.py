import pandas as pd
import os
import re
import time
import shutil
from datetime import datetime
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, InvalidSessionIdException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

# Path to Excel file (in parent folder)
excel_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "ScrappedProducts.xlsx")
backup_folder = os.path.join(os.path.dirname(os.path.dirname(__file__)), "Backups")
MAX_BACKUPS = 5  # Keep last 5 backups

# Read the Excel file
print("Reading Excel file...")
df = pd.read_excel(excel_path)

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
        
        driver = webdriver.Chrome(options=chrome_options)
        driver.implicitly_wait(3)  # Reduced from 10 to 3 seconds for faster failure detection
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
        driver.get(link)
        
        # ULTRA-FAST DETECTION: Minimal wait, then immediately check for product
        try:
            # Wait only for body (very fast)
            WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        except:
            print(f"    Product not found (page load timeout)")
            return "Product not found", "Product not found", "Product not found", None
        
        # IMMEDIATE CHECK: Look for key product elements (fastest way to detect if product exists)
        # If we can't find the unit element quickly, product likely doesn't exist
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
        
        # Try to find the unit element with a very short timeout (1 second)
        # This is the fastest way to confirm product exists
        try:
            uom_element = WebDriverWait(driver, 1).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "span.ess-detail-uom, .ess-detail-uom"))
            )
            # If we found it, product exists - continue to extraction
        except:
            # Unit element not found - do quick checks to confirm product doesn't exist
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
                # Fallback: Try XPath
                try:
                    uom_elements = driver.find_elements(By.XPATH, "//*[contains(@class, 'ess-detail-uom') or contains(@class, 'ess-product-uom')]")
                    for elem in uom_elements:
                        text = elem.text.strip()
                        if text.startswith('/'):
                            unit = text[1:].strip().upper()
                            if unit and len(unit) >= 2 and len(unit) <= 4:
                                website_unit = unit
                                print(f"    Found unit (XPath): {website_unit}")
                                break
                except:
                    pass
            
            # If still not found, try page source regex (quick check)
            if not website_unit:
                page_source = driver.page_source
                uom_pattern = r'class="ess-detail-uom"[^>]*>/([A-Z]{2,4})\b'
                match = re.search(uom_pattern, page_source, re.IGNORECASE)
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
            print(f"    Unit mismatch! Skipping...")
            return "Unit not matched", "Unit not matched", "Unit not matched", None
        
        # Unit matches, proceed with scraping
        print(f"    Unit matched! Scraping data...")
        
        # 1. Scrape Image URL - FIXED to only get actual product images
        image_url = None
        try:
            # Define image file extensions and non-image extensions to exclude
            image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.webp', '.bmp', '.svg']
            non_image_extensions = ['.js', '.css', '.html', '.xml', '.json', '.pdf', '.txt', '.ico']
            
            # Helper function to check if URL is an image
            def is_image_url(url):
                if not url:
                    return False
                url_lower = url.lower()
                # Exclude non-image extensions
                if any(url_lower.endswith(ext) for ext in non_image_extensions):
                    return False
                # Prefer URLs with image extensions
                if any(url_lower.endswith(ext) for ext in image_extensions):
                    return True
                # If from known image domains, accept it
                if any(domain in url_lower for domain in ['oppictures.com', 'content.oppictures.com']):
                    return True
                # Exclude script and style sources
                if any(exclude in url_lower for exclude in ['bazaarvoice', 'script', 'style', 'analytics', 'tracking']):
                    return False
                return False  # Default: be conservative, only accept if we're sure
            
            # METHOD 1: Use DOM to find actual <img> tags (most reliable)
            try:
                # First, look specifically for oppictures images (most reliable)
                img_elements = driver.find_elements(By.XPATH, "//img[contains(@src, 'oppictures')]")
                for img in img_elements:
                    src = img.get_attribute('src')
                    if src and is_image_url(src):
                        if src.startswith('//'):
                            src = 'https:' + src
                        image_url = src
                        print(f"    Found image from oppictures img tag: {image_url[:80]}...")
                        break
                
                # If not found, look for main product image (usually the largest or first one)
                if not image_url:
                    # Look for images in product detail sections
                    product_img_selectors = [
                        "//img[contains(@class, 'product') or contains(@class, 'item')]",
                        "//img[contains(@id, 'product') or contains(@id, 'item')]",
                        "//div[contains(@class, 'product')]//img",
                        "//div[contains(@class, 'detail')]//img",
                    ]
                    
                    for selector in product_img_selectors:
                        try:
                            img_elements = driver.find_elements(By.XPATH, selector)
                            for img in img_elements:
                                src = img.get_attribute('src')
                                if src and is_image_url(src):
                                    if src.startswith('//'):
                                        src = 'https:' + src
                                    # Skip small images (likely icons/logos)
                                    try:
                                        width = img.get_attribute('width')
                                        height = img.get_attribute('height')
                                        if width and height:
                                            w, h = int(width), int(height)
                                            if w < 50 or h < 50:  # Skip very small images
                                                continue
                                    except:
                                        pass
                                    image_url = src
                                    print(f"    Found image from product img tag: {image_url[:80]}...")
                                    break
                            if image_url:
                                break
                        except:
                            continue
            except Exception as e:
                print(f"    Error in DOM image search: {e}")
            
            # METHOD 2: Fallback to page source regex (only for <img> tags)
            if not image_url:
                # Look specifically for <img> tags with oppictures
                img_tag_pattern = r'<img[^>]+src=["\']([^"\']*oppictures[^"\']+)["\']'
                matches = re.finditer(img_tag_pattern, page_source, re.IGNORECASE)
                for match in matches:
                    src = match.group(1)
                    if is_image_url(src):
                        if src.startswith('//'):
                            src = 'https:' + src
                        image_url = src
                        print(f"    Found image from page source (oppictures): {image_url[:80]}...")
                        break
                
                # If still not found, look for any <img> tag with image extensions
                if not image_url:
                    img_tag_pattern = r'<img[^>]+src=["\']([^"\']+)["\']'
                    matches = re.finditer(img_tag_pattern, page_source, re.IGNORECASE)
                    for match in matches:
                        src = match.group(1)
                        if is_image_url(src):
                            # Skip common non-product images
                            if any(skip in src.lower() for skip in ['logo', 'icon', 'banner', 'header', 'footer', 'button', 'arrow']):
                                continue
                            if src.startswith('//'):
                                src = 'https:' + src
                            image_url = src
                            print(f"    Found image from page source: {image_url[:80]}...")
                            break
                                
        except Exception as e:
            print(f"    Error finding image: {e}")
        
        # 2. Scrape Product Name (Global Product Type from Product Details section)
        # Learned pattern: Product name is in <td> element that follows "Global Product Type"
        product_name = None
        try:
            # METHOD 1: Use specific XPath (learned from inspection - most reliable)
            try:
                # Find the td containing "Global Product Type" and get the following sibling td
                name_element = driver.find_element(By.XPATH, "//td[contains(text(), 'Global Product Type')]/following-sibling::td[1]")
                product_name = name_element.text.strip()
                if product_name:
                    print(f"    Found product name from td: {product_name}")
            except:
                # Alternative: Find all tds in the same row and get the one after "Global Product Type"
                try:
                    name_elements = driver.find_elements(By.XPATH, "//td[contains(text(), 'Global Product Type')]/../td")
                    for elem in name_elements:
                        text = elem.text.strip()
                        if text and text != 'Global Product Type' and len(text) > 5:
                            product_name = text
                            print(f"    Found product name from alternative td: {product_name}")
                            break
                except:
                    pass
            
            # METHOD 2: Fallback to page source pattern
            if not product_name:
                pattern = r'Global Product Type[:\s]+([^\n<]+)'
                match = re.search(pattern, page_source, re.IGNORECASE)
                if match:
                    product_name = match.group(1).strip()
                    product_name = re.sub(r'<[^>]+>', '', product_name).strip()
                    if product_name:
                        print(f"    Found product name from page source: {product_name}")
                    
        except Exception as e:
            print(f"    Error finding product name: {e}")
        
        # 3. Scrape Description
        # Learned pattern: Description is in a <div> element
        # Must exclude: "Non-Stock Item - Extended Delivery Time" and similar stock/status messages
        description = None
        exclude_patterns = [
            'non-stock item',
            'extended delivery time',
            'stock item',
            'delivery time',
            'description:',  # Just the label
            'in stock',
            'out of stock',
        ]
        
        try:
            # METHOD 1: Find div containing actual description text (learned from inspection)
            # The description appears in a div element, but we need to skip stock messages
            try:
                # Look for divs with substantial text content that looks like a product description
                all_divs = driver.find_elements(By.TAG_NAME, "div")
                for div in all_divs:
                    text = div.text.strip()
                    # Look for description-like content (substantial text, reasonable length)
                    if text and 30 < len(text) < 500:  # Reasonable description length
                        text_lower = text.lower()
                        # Exclude stock messages, navigation, labels
                        if any(exclude in text_lower for exclude in exclude_patterns):
                            continue
                        if any(skip in text_lower for skip in ['menu', 'nav', 'header', 'footer', 'cookie', 'terms', 'privacy']):
                            continue
                        
                        # Check if it contains common description words (actual product description)
                        if any(word in text_lower for word in ['for', 'with', 'set', 'includes', 'features', 'replacement', 'key', 'lock']):
                            description = text
                            print(f"    Found description from div: {description[:80]}...")
                            break
            except Exception as e:
                print(f"    Error in div search: {e}")
            
            # METHOD 2: Try elements with description class/id (but filter out stock messages)
            if not description:
                desc_selectors = [
                    "//div[contains(@class, 'description')]",
                    "//div[contains(@id, 'description')]",
                ]
                for selector in desc_selectors:
                    try:
                        desc_elements = driver.find_elements(By.XPATH, selector)
                        for elem in desc_elements:
                            text = elem.text.strip()
                            if text and len(text) > 30:
                                text_lower = text.lower()
                                # Skip if it's a stock message or label
                                if not any(exclude in text_lower for exclude in exclude_patterns):
                                    description = text
                                    print(f"    Found description from class/id: {description[:80]}...")
                                    break
                        if description:
                            break
                    except:
                        continue
            
            # METHOD 3: Fallback to page source pattern (filtered)
            if not description:
                desc_pattern = r'<[^>]*class="[^"]*description[^"]*"[^>]*>([^<]+(?:<[^>]+>[^<]+)*)</[^>]*>'
                matches = list(re.finditer(desc_pattern, page_source, re.IGNORECASE | re.DOTALL))
                for match in matches:
                    desc_text = match.group(1)
                    desc_text = re.sub(r'<[^>]+>', ' ', desc_text)
                    desc_text = ' '.join(desc_text.split())
                    if desc_text and len(desc_text) > 20:
                        desc_lower = desc_text.lower()
                        # Skip if it's a stock message
                        if not any(exclude in desc_lower for exclude in exclude_patterns):
                            description = desc_text
                            print(f"    Found description from page source: {description[:80]}...")
                            break
                    
        except Exception as e:
            print(f"    Error finding description: {e}")
        
        return product_name, description, image_url, website_unit
        
    except TimeoutException:
        print(f"    Timeout loading page")
        return None, None, None, None
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
    
    for idx in range(start_idx, end_idx + 1):
        row = df.iloc[idx]
        print(f"Row {idx + 1}/{len(df)} (Processing {processed_count + 1}/{total_to_process}):")
        
        # Get link
        link = row[link_col]
        if pd.isna(link) or str(link).strip() == '':
            print("  No link found, skipping...")
            skipped_count += 1
            continue
        
        # Get expected unit
        expected_unit = row[unit_col]
        if pd.isna(expected_unit):
            print("  No unit of measure found, skipping...")
            skipped_count += 1
            continue
        
        # Check if already scraped (skip if Product Name is already filled)
        if pd.notna(row[product_name_col]) and str(row[product_name_col]).strip() != '':
            product_name_value = str(row[product_name_col]).strip()
            # Skip if already has a real product name (not error messages)
            if product_name_value not in ['Unit not matched', 'Product not found']:
                print("  Already scraped, skipping...")
                skipped_count += 1
                continue
            # Skip if already marked as "Product not found" (to avoid re-checking unavailable products)
            elif product_name_value == 'Product not found':
                print("  Already marked as 'Product not found', skipping...")
                skipped_count += 1
                continue
            # Skip if already marked as "Unit not matched" (to avoid re-checking products with wrong units)
            elif product_name_value == 'Unit not matched':
                print("  Already marked as 'Unit not matched', skipping...")
                skipped_count += 1
                continue
        
        # Scrape data (with retry on session loss)
        max_retries = 2
        retry_count = 0
        product_name, description, image_url, website_unit = None, None, None, None
        
        while retry_count <= max_retries:
            product_name, description, image_url, website_unit = scrape_product_data(str(link).strip(), expected_unit)
            
            # If we got results (even if error), break
            if product_name is not None or retry_count >= max_retries:
                break
            
            # If we got None, None, None, None and it might be a session issue, retry
            retry_count += 1
            if retry_count <= max_retries:
                print(f"    Retrying ({retry_count}/{max_retries})...")
                time.sleep(1)  # Brief pause before retry
        
        # Update Excel row
        if product_name == "Unit not matched":
            df.at[idx, product_name_col] = "Unit not matched"
            df.at[idx, description_col] = "Unit not matched"
            df.at[idx, image_url_col] = "Unit not matched"
            processed_count += 1
        elif product_name == "Product not found":
            df.at[idx, product_name_col] = "Product not found"
            df.at[idx, description_col] = "Product not found"
            df.at[idx, image_url_col] = "Product not found"
            processed_count += 1
        elif product_name:
            df.at[idx, product_name_col] = product_name
            if description:
                df.at[idx, description_col] = description
            if image_url:
                df.at[idx, image_url_col] = image_url
            processed_count += 1
        else:
            error_count += 1
        
        # Save progress (every 10 products)
        if processed_count % 10 == 0 and processed_count > 0:
            print(f"\nSaving progress...")
            df.to_excel(excel_path, index=False)
            print(f"Progress saved! (Processed: {processed_count}, Skipped: {skipped_count}, Errors: {error_count})\n")
        
        # Small delay to avoid overwhelming the server (reduced)
        time.sleep(0.5)
    
    # Final save
    print(f"\nSaving final results...")
    df.to_excel(excel_path, index=False)
    
    # Create backup after completion
    print("Creating backup after completion...")
    create_backup()
    
    print(f"\nDone! Processed: {processed_count}, Skipped: {skipped_count}, Errors: {error_count}")

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

