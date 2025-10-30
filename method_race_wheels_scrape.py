from seleniumbase import Driver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException
import logging
import tempfile
import psutil
import shutil
import os
import time
import pandas as pd
import xlwings as xw
from datetime import datetime


# Constants 
WEBSITE_URL = "https://www.methodracewheels.com/"
CAPTCHA_WAIT_TIME = 500
ELEMENT_WAIT_TIME = 50
PAGE_LOAD_WAIT_TIME = 50

# Logging configuration
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s: %(message)s',
    handlers=[
        logging.FileHandler('methodracewheels_automation.log'),
        logging.StreamHandler()
    ]
)

# Setup Selenium WebDriver with temporary user data directory
def setup_driver():
    try:
        user_data_dir = tempfile.mkdtemp()
        driver = Driver(uc=True)
        driver.user_data_dir = user_data_dir
        driver.set_page_load_timeout(30)
        return driver
    except Exception as e:
        logging.error(f"Driver setup failed: {e}")
        raise

# Wait for user to solve CAPTCHA
def wait_for_captcha(driver):
    print("Please solve the CAPTCHA manually.")
    input("Press Enter after solving the CAPTCHA...")

    # Ensure page is ready after CAPTCHA
    WebDriverWait(driver, PAGE_LOAD_WAIT_TIME).until(
        lambda d: d.execute_script('return document.readyState') == 'complete'
    )
    print("CAPTCHA solved and page loaded successfully.")

# Scrape product links from the main collection page
def scrape_product_links(driver):
    product_links = []
    
    try:
        print("Scraping Method Race Wheels...")
        driver.get("https://www.methodracewheels.com/collections/standard-wheels")
        time.sleep(5)

        # Scroll to load all products
        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        # Find product links using multiple possible selectors
        selectors = [
            "a[href*='/products/']",
            ".product-card a",
            "a.product-item-link",
            "a[href*='/collections/'][href*='/products/']"
        ]
        
        for selector in selectors:
            try:
                elements = driver.find_elements(By.CSS_SELECTOR, selector)
                if elements:
                    for element in elements:
                        href = element.get_attribute("href")
                        if href and "/products/" in href and href not in product_links:
                            product_links.append(href)
                    break
            except:
                continue

        print(f"✅ Found {len(product_links)} unique product links.")
        print("Product links:", product_links)

    except Exception as e:
        print(f"❌ Error scraping: {str(e)}")
    

    return product_links

# parse product details from individual product pages
def parse_product_details(driver, url):
    product_data = {}  # Removed 'URL' from initialization
    
    try:
        driver.get(url)
        time.sleep(5)  # Increased wait time for dynamic content
        
        # Parse overview from meta description
        try:
            description = driver.find_element(By.CSS_SELECTOR, 'meta[itemprop="description"]').get_attribute('content')
            product_data['Overview'] = description
        except:
            product_data['Overview'] = ''
        
        # Parse Details section - try multiple approaches
        try:
            print("Looking for Details section...")
            
            # Try different heading selectors
            heading_selectors = [
                'h2.product_details-title',
                'h2[class*="product"]',
                'h2[class*="details"]',
                '.rte h2',
                'h2'
            ]
            
            headings = []
            for selector in heading_selectors:
                headings = driver.find_elements(By.CSS_SELECTOR, selector)
                if headings:
                    print(f"Found {len(headings)} headings with selector: {selector}")
                    break
            
            if headings:
                for i, heading in enumerate(headings):
                    heading_text = heading.text.strip()
                    print(f"Heading {i}: '{heading_text}'")
                    
                    if heading_text.lower() == 'details':
                        print("Found Details heading, looking for list items...")
                        
                        # Try to find the next sibling or parent container
                        try:
                            # Method 1: Get parent and find ul
                            parent = heading.find_element(By.XPATH, '..')
                            detail_items = parent.find_elements(By.CSS_SELECTOR, 'ul li')
                            
                            if not detail_items:
                                # Method 2: Get following sibling div
                                following_div = heading.find_element(By.XPATH, 'following-sibling::div[1]')
                                detail_items = following_div.find_elements(By.TAG_NAME, 'li')
                            
                            if not detail_items:
                                # Method 3: Search in parent's parent
                                grandparent = heading.find_element(By.XPATH, '../..')
                                detail_items = grandparent.find_elements(By.CSS_SELECTOR, 'ul li')
                            
                            print(f"Found {len(detail_items)} detail items")
                            
                            if detail_items:
                                bullet_number = 1
                                for j, item in enumerate(detail_items):
                                    text = item.text.strip()
                                    if text:
                                        # Skip if it looks like a specification (has spans)
                                        spans = item.find_elements(By.TAG_NAME, 'span')
                                        if len(spans) >= 2:
                                            print(f"Skipping item {j} - looks like a spec")
                                            continue
                                        
                                        print(f"Detail item {j}: '{text[:50]}...'")
                                        # Clean up the text
                                        text = text.replace('(1bs)', '(lbs)').replace('(Ibs)', '(lbs)')
                                        
                                        # Create separate column for each bullet
                                        column_name = f'Bullet {bullet_number}'
                                        product_data[column_name] = text
                                        print(f"✓ Added {column_name}")
                                        bullet_number += 1
                                
                                if bullet_number > 1:
                                    print(f"✓ Added {bullet_number - 1} bullet point columns")
                                    break
                        except Exception as e:
                            print(f"Error extracting details from heading: {e}")
                            continue
                    
        except Exception as e:
            print(f"Error parsing Details section: {e}")
            import traceback
            traceback.print_exc()
        
        # Parse specifications table dynamically
        print(f"Checking specifications structure...")
        
        # Find Specifications section specifically
        try:
            spec_items = []
            
            # First try to find Specifications heading
            if headings:
                for heading in headings:
                    if heading.text.strip().lower() == 'specifications':
                        parent = heading.find_element(By.XPATH, '..')
                        spec_items = parent.find_elements(By.CSS_SELECTOR, 'ul.product_specs-list li')
                        if not spec_items:
                            spec_items = parent.find_elements(By.CSS_SELECTOR, 'ul li')
                        print(f"Found {len(spec_items)} specification items in Specifications section")
                        break
            
            if not spec_items:
                # Fallback to original selectors
                selectors_to_try = [
                    '.product_specs-list li',
                    '.product_details-text ul li',
                    'ul.product_specs-list li',
                    '.product_details li',
                    '.rte ul li',
                    'div[class*="spec"] li',
                    'div[class*="product_details"] li'
                ]
                
                for selector in selectors_to_try:
                    spec_items = driver.find_elements(By.CSS_SELECTOR, selector)
                    if spec_items:
                        print(f"Found {len(spec_items)} items with selector: {selector}")
                        break
            
            if not spec_items:
                print(f"No specifications found with any selector")
            else:
                print(f"Processing {len(spec_items)} specification items...")
                
                for item in spec_items:
                    try:
                        spans = item.find_elements(By.TAG_NAME, 'span')
                        if len(spans) >= 2:
                            key = spans[0].text.strip()
                            value = spans[1].text.strip()
                            
                            if key and value:
                                # Clean up common typos
                                key = key.replace('(1bs)', '(lbs)').replace('(Ibs)', '(lbs)')
                                product_data[key] = value
                                print(f"Added: '{key}' = '{value}'")
                    except Exception as e:
                        print(f"Error parsing spec item: {e}")
                        continue
                        
        except Exception as e:
            print(f"Error parsing specifications: {e}")
        
        # Reorder columns according to specified arrangement
        ordered_data = {}
        
        # Define the desired column order
        priority_columns = [
            'Part Number',
            'Overview',
            'Wheel Diameter (in)',
            'Wheel Width (in)',
            'Bolt Pattern',
            'Offset (mm)',
            'Hub Bore (mm)',
            'Back Spacing (in)',
            'Wheel Weight (lbs)',
            'Max Load (lbs)'
        ]
        
        # Add priority columns first (if they exist)
        for col in priority_columns:
            if col in product_data:
                ordered_data[col] = product_data[col]
        
        # Add all Bullet columns in order
        bullet_cols = sorted([k for k in product_data.keys() if k.startswith('Bullet ')], 
                            key=lambda x: int(x.split()[1]))
        for col in bullet_cols:
            ordered_data[col] = product_data[col]
        
        # Add any remaining columns that weren't in priority list or bullets
        for col in product_data:
            if col not in ordered_data:
                ordered_data[col] = product_data[col]
        
        return ordered_data
                        
    except Exception as e:
        print(f"Error parsing {url}: {e}")
        return {}


# Save data to Excel with specific column ordering and autofit
def save_to_excel(data, desktop_path):
    if not data:
        print("No data to save")
        return
    
    # Create filename with current date
    current_date = datetime.now().strftime("%Y_%m_%d")
    filename = f"method_race_wheels_sample_scrape_data_{current_date}.xlsx"
    filepath = os.path.join(desktop_path, filename)
    
    # Get all unique columns from all products
    all_columns = set()
    for product in data:
        all_columns.update(product.keys())
    
    # Define specific column order
    preferred_order = [
        'Overview', 'Part Number', 'Wheel Diameter (in)', 'Wheel Width (in)', 
        'Bolt Pattern', 'Offset (mm)', 'Hub Bore (mm)', 'Back Spacing (in)', 
        'Wheel Weight (lbs)', 'Max Load (lbs)'
    ]
    
    # Separate bullet columns and other columns
    bullet_columns = [col for col in all_columns if col.startswith('Bullet ')]
    other_columns = [col for col in all_columns if col not in preferred_order and not col.startswith('Bullet ')]
    
    # Sort bullet columns numerically
    bullet_columns.sort(key=lambda x: int(x.split(' ')[1]) if x.split(' ')[1].isdigit() else 0)
    
    # Sort other remaining columns alphabetically
    other_columns.sort()
    
    # Combine all columns in order
    columns = [col for col in preferred_order if col in all_columns] + bullet_columns + other_columns
    
    print(f"Creating Excel with columns: {columns}")
    
    # Create DataFrame with ordered columns and remove duplicates
    df = pd.DataFrame(data, columns=columns)
    df = df.drop_duplicates()
    print(f"Removed duplicates: {len(data)} -> {len(df)} rows")
    df.to_excel(filepath, index=False)
    
    # Use xlwings to autofit columns
    try:
        app = xw.App(visible=False)
        wb = app.books.open(filepath)
        ws = wb.sheets[0]
        ws.autofit()
        wb.save()
        wb.close()
        app.quit()
        print(f"✅ Excel file saved: {filepath}")
    except Exception as e:
        print(f"Error with xlwings autofit: {e}")

# Cleanup function to close driver and remove temp files
def cleanup(driver):
    try:
        try:
            driver.quit()
        except Exception as e:
            print(f"Driver quit error (can be ignored): {e}")
        
        time.sleep(2)
        
        # Clean up any remaining chrome processes
        try:
            for proc in psutil.process_iter(attrs=['pid', 'name']):
                try:
                    if 'chrome' in proc.info['name'].lower():
                        proc.kill()
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
        except Exception as e:
            print(f"Process cleanup error (can be ignored): {e}")
        
        # Remove temporary directory
        if hasattr(driver, 'user_data_dir') and os.path.exists(driver.user_data_dir):
            try:
                shutil.rmtree(driver.user_data_dir, ignore_errors=True)
            except (FileNotFoundError, PermissionError) as e:
                print(f"Temp directory cleanup error (can be ignored): {e}")
    except Exception as close_error:
        print(f"Error during cleanup (can be ignored): {close_error}")

# Main execution flow
def main():
    driver = setup_driver()
    try:
        logging.info("Navigating to the website.")
        driver.get(WEBSITE_URL)

        # Wait for CAPTCHA and ensure page readiness
        wait_for_captcha(driver)

        print("Collecting part links...")
        part_links = scrape_product_links(driver)
        
        # Parse details for each product
        all_products = []
        for i, link in enumerate(part_links):
            print(f"Parsing product {i+1}/{len(part_links)}: {link}")
            product_data = parse_product_details(driver, link)
            if product_data:
                all_products.append(product_data)
                print(f"Parsed: {len(product_data)} fields - {list(product_data.keys())}")
                all_products.append(product_data)
        
        print(f"✅ Scraping complete. {len(part_links)} links found, {len(all_products)} products parsed.")
        
        # Save to Excel on desktop
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        save_to_excel(all_products, desktop_path)
    except Exception as e:
        logging.error(f"Failed to load the website: {e}")
        driver.quit()
    finally:
        if driver:
            cleanup(driver)
        logging.info("Driver closed.")

# Run the main function
if __name__ == "__main__":
    main()
