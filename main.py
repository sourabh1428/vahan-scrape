from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import requests
import logging
import random

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# First verify Indian IP requirement
def check_indian_ip():
    try:
        response = requests.get('https://api.ipify.org?format=json', timeout=5)
        ip = response.json()['ip']
        ip_data = requests.get(f"http://ip-api.com/json/{ip}", timeout=5).json()
        if ip_data.get('countryCode') != 'IN':
            logger.error(f"Error: Non-Indian IP detected. Connect to Indian VPN first.")
            logger.error(f"Current IP: {ip_data.get('query')} ({ip_data.get('country')})")
            return False
        logger.info(f"IP check passed: {ip_data.get('query')} ({ip_data.get('country')})")
        return True
    except Exception as e:
        logger.warning(f"IP check failed with error: {str(e)}. Proceeding anyway...")
        return True  # Proceed if IP check fails

# Random delay to simulate human behavior
def random_delay(min_seconds=1, max_seconds=3):
    delay = random.uniform(min_seconds, max_seconds)
    time.sleep(delay)
    return delay

# Create downloads directory if it doesn't exist
download_dir = os.path.join(os.getcwd(), 'downloads')
os.makedirs(download_dir, exist_ok=True)
logger.info(f"Download directory: {download_dir}")

if not check_indian_ip():
    exit()

# Configure Chrome with advanced options
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option("useAutomationExtension", False)
chrome_options.add_argument("--disable-web-security")
chrome_options.add_argument("--allow-running-insecure-content")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-popup-blocking")

# Add a real user agent
user_agents = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.51 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.75 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/15.4 Safari/605.1.15"
]
chrome_options.add_argument(f"--user-agent={random.choice(user_agents)}")

prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "profile.default_content_setting_values.geolocation": 1  # Allow geolocation
}
chrome_options.add_experimental_option("prefs", prefs)

try:
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    logger.info("WebDriver initialized successfully")

    # Bypass initial security checks
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
        Object.defineProperty(navigator, 'webdriver', {
            get: () => undefined
        });
        window.navigator.chrome = {
            runtime: {},
        };
        const originalQuery = window.navigator.permissions.query;
        window.navigator.permissions.query = (parameters) => (
            parameters.name === 'notifications' ? 
                Promise.resolve({ state: Notification.permission }) :
                originalQuery(parameters)
        );
        """
    })
    
    # Load page and wait for it to be fully loaded
    logger.info("Loading the Vahan website...")
    driver.get("https://vahan.parivahan.gov.in/vahan4dashboard/vahan/view/reportview.xhtml")
    
    # Wait for page to fully load
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "[name='javax.faces.ViewState']"))
    )
    logger.info("Page loaded successfully")
    
    # Let the page fully render with a realistic delay
    delay = random_delay(3, 6)
    logger.info(f"Waiting {delay:.2f} seconds for page to stabilize")
    
    # Step 1: Select a state from the dropdown
    logger.info("Selecting a specific state...")
    try:
        # Click on the state dropdown to open it
        state_dropdown = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "j_idt37_label"))
        )
        state_dropdown.click()
        logger.info("Clicked on state dropdown")
        random_delay(1, 2)
        
        # Select a state (e.g., "Maharashtra")
        # You can change this to any state you prefer
        state_to_select = "Karnataka(68)"
        state_option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f"//li[contains(@data-label, '{state_to_select}')]"))
        )
        state_option.click()
        logger.info(f"Selected state: {state_to_select}")
        random_delay(1, 2)
    except Exception as e:
        logger.warning(f"Error selecting state: {str(e)}")
        # Try an alternative approach using JavaScript
        try:
            driver.execute_script("PrimeFaces.widgets.widget_j_idt37.selectValue('20');")  # Index for Maharashtra
            logger.info("Selected state using JavaScript")
        except Exception as js_e:
            logger.warning(f"JavaScript state selection failed: {str(js_e)}")
    
    # Step 2: Select 'Maker' for Y-axis
    logger.info("Selecting 'Maker' for Y-axis...")
    try:
        # Click on the Y-axis dropdown to open it
        y_axis_dropdown = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "yaxisVar_label"))
        )
        y_axis_dropdown.click()
        logger.info("Clicked on Y-axis dropdown")
        random_delay(1, 2)
        
        # Select 'Maker' option
        maker_option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//li[@data-label='Maker']"))
        )
        maker_option.click()
        logger.info("Selected 'Maker' for Y-axis")
        random_delay(1, 2)
    except Exception as e:
        logger.warning(f"Error selecting Y-axis: {str(e)}")
        # Try an alternative approach using JavaScript
        try:
            driver.execute_script("PrimeFaces.widgets.widget_yaxisVar.selectValue('4');")  # Index for Maker
            logger.info("Selected 'Maker' for Y-axis using JavaScript")
        except Exception as js_e:
            logger.warning(f"JavaScript Y-axis selection failed: {str(js_e)}")
    
    # Step 3: Select 'Month Wise' for X-axis
    logger.info("Selecting 'Month Wise' for X-axis...")
    try:
        # Click on the X-axis dropdown to open it
        x_axis_dropdown = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "xaxisVar_label"))
        )
        x_axis_dropdown.click()
        logger.info("Clicked on X-axis dropdown")
        random_delay(1, 2)
        
        # Select 'Month Wise' option
        month_wise_option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//li[@data-label='Month Wise']"))
        )
        month_wise_option.click()
        logger.info("Selected 'Month Wise' for X-axis")
        random_delay(1, 2)
    except Exception as e:
        logger.warning(f"Error selecting X-axis: {str(e)}")
        # Try an alternative approach using JavaScript
        try:
            driver.execute_script("PrimeFaces.widgets.widget_xaxisVar.selectValue('6');")  # Index for Month Wise
            logger.info("Selected 'Month Wise' for X-axis using JavaScript")
        except Exception as js_e:
            logger.warning(f"JavaScript X-axis selection failed: {str(js_e)}")
    
    # Step 4: Click the first refresh button on the right side
    logger.info("Clicking the first Refresh button on the right side...")
    try:
        refresh_button_right = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "j_idt73"))
        )
        refresh_button_right.click()
        logger.info("Clicked first Refresh button")
        
        # Wait for data to load
        random_delay(5, 8)
    except Exception as e:
        logger.warning(f"Error clicking first refresh button: {str(e)}")
        # Try an alternative approach
        try:
            # Find all buttons and try to find the refresh button
            buttons = driver.find_elements(By.TAG_NAME, "button")
            for button in buttons:
                if "refresh" in button.get_attribute("id").lower() or "refr" in button.get_attribute("id").lower():
                    button.click()
                    logger.info(f"Clicked alternative refresh button: {button.get_attribute('id')}")
                    random_delay(5, 8)
                    break
        except Exception as btn_e:
            logger.warning(f"Alternative refresh button approach failed: {str(btn_e)}")
    
    # Step 5: Expand the vehicle category filter if needed
    logger.info("Checking if vehicle category filter needs to be expanded...")
    try:
        # Check if the filter is collapsed
        toggle_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "filterLayout-toggler"))
        )
        
        # Check if it's in closed state
        if "ui-layout-toggler-closed" in toggle_element.get_attribute("class"):
            # Click on the toggle element to expand
            toggle_element.click()
            logger.info("Expanded the vehicle category filter")
            random_delay(2, 3)
        else:
            logger.info("Vehicle category filter is already expanded")
        
    except Exception as e:
        logger.warning(f"Error expanding vehicle category filter: {str(e)}")
    
    # Step 6: Select all required options
    logger.info("Selecting all required options...")
    try:
        # Wait for the vehicle category table to be present and visible
        vehicle_table = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, 'VhCatg'))
        )
        
        # Find all two-wheeler labels
        two_wheeler_labels = vehicle_table.find_elements(By.XPATH, ".//label[contains(text(), 'TWO WHEELER')]")
        
        # Find electric vehicle labels
        electric_vehicle_labels = vehicle_table.find_elements(By.XPATH, ".//label[text()='PURE EV' or text()='ELECTRIC(BOV)']")
        
        # Combine both lists
        all_labels = two_wheeler_labels + electric_vehicle_labels
        
        for label in all_labels:
            checkbox_id = label.get_attribute("for")
            checkbox = driver.find_element(By.ID, checkbox_id)
            
            # Only click if not already selected
            if not checkbox.is_selected():
                # Scroll to the checkbox to ensure it's visible
                driver.execute_script("arguments[0].scrollIntoView(true);", checkbox)
                random_delay(0.5, 1)
                
                # Click the checkbox
                checkbox.click()
                logger.info(f"Selected checkbox: {label.text}")
                
                # Wait for the checkbox to be checked
                WebDriverWait(driver, 5).until(
                    EC.element_to_be_selected(checkbox)
                )
                logger.info(f"Checkbox {label.text} is now selected")
            
            random_delay(0.5, 1)
        
        logger.info("All required options selected")
        random_delay(2, 3)
        
    except Exception as e:
        logger.warning(f"Error selecting options: {str(e)}")
        try:
            # Alternative approach using JavaScript if regular clicks fail
            logger.info("Attempting JavaScript approach to select checkboxes...")
            # Select two-wheeler options by their values
            two_wheeler_checkboxes = driver.find_elements(By.XPATH, "//input[@name='VhCatg' and (@value='2WIC' or @value='2WN' or @value='2WT')]")
            # Select electric vehicle options by their values
            electric_vehicle_checkboxes = driver.find_elements(By.XPATH, "//input[@name='fuel' and (@value='4' or @value='22')]")
            
            # Combine both lists
            all_checkboxes = two_wheeler_checkboxes + electric_vehicle_checkboxes
            
            for checkbox in all_checkboxes:
                if not checkbox.is_selected():
                    driver.execute_script("arguments[0].click();", checkbox)
                    logger.info(f"Selected checkbox via JavaScript: {checkbox.get_attribute('id')}")
                    random_delay(0.5, 1)
        except Exception as js_e:
            logger.warning(f"JavaScript selection failed: {str(js_e)}")
    
    # Step 7: Click the second refresh button on the left side
    logger.info("Clicking the second Refresh button on the left side...")
    try:
        refresh_button_left = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "j_idt82"))
        )
        refresh_button_left.click()
        logger.info("Clicked second Refresh button")
        
        # Wait for data to load
        random_delay(5, 8)
    except Exception as e:
        logger.warning(f"Error clicking second refresh button: {str(e)}")
        # Try an alternative approach
        try:
            # Find all buttons and try to find the refresh button
            buttons = driver.find_elements(By.TAG_NAME, "button")
            for button in buttons:
                if "refresh" in button.get_attribute("id").lower() or "refr" in button.get_attribute("id").lower():
                    button.click()
                    logger.info(f"Clicked alternative refresh button: {button.get_attribute('id')}")
                    random_delay(5, 8)
                    break
        except Exception as btn_e:
            logger.warning(f"Alternative refresh button approach failed: {str(btn_e)}")
    
    # Step 8: Download the Excel file
    logger.info("Attempting to download Excel file...")
    try:
        # Wait for a moment to ensure the data is fully loaded
        random_delay(3, 5)
        
        # Look for Excel download button
        excel_button_locators = [
            (By.ID, "vchgroupTable:xls"),
            (By.ID, "groupingTable:j_idt204"),
            (By.CSS_SELECTOR, "button[id*='xls'], a[id*='xls'], span[id*='xls']"),
            (By.XPATH, "//button[contains(@id, 'xls')]"),
            (By.XPATH, "//button[contains(@title, 'Excel') or contains(@title, 'XLS')]")
        ]
        
        excel_button = None
        for locator_type, locator in excel_button_locators:
            try:
                excel_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((locator_type, locator))
                )
                if excel_button:
                    break
            except:
                continue
        
        if excel_button:
            # Scroll to the button to make it visible
            driver.execute_script("arguments[0].scrollIntoView(true);", excel_button)
            random_delay(1, 2)
            
            # Click the Excel button
            excel_button.click()
            logger.info("Clicked Excel download button")
            
            # Wait for download to start
            random_delay(5, 8)
        else:
            logger.warning("Excel button not found using standard locators")
            
            # Try to find any button with an Excel icon or title
            buttons = driver.find_elements(By.TAG_NAME, "button")
            for button in buttons:
                button_html = button.get_attribute("outerHTML")
                if "excel" in button_html.lower() or "xls" in button_html.lower():
                    logger.info(f"Found potential Excel button: {button.get_attribute('id')}")
                    driver.execute_script("arguments[0].scrollIntoView(true);", button)
                    random_delay(1, 2)
                    button.click()
                    logger.info("Clicked potential Excel button")
                    random_delay(5, 8)
                    break
            
            # If still not found, try JavaScript approach
            logger.info("Attempting JavaScript approach for Excel download")
            js_approaches = [
                "PrimeFaces.widgets.widget_groupingTable.exportExcel();",
                "document.querySelector('button[id*=\"xls\"]').click();",
                "document.querySelector('a[id*=\"xls\"]').click();"
            ]
            
            for js in js_approaches:
                try:
                    driver.execute_script(js)
                    logger.info(f"Executed JavaScript: {js}")
                    random_delay(5, 8)
                except Exception as js_error:
                    logger.warning(f"JavaScript failed: {str(js_error)}")
    
    except Exception as e:
        logger.warning(f"Error downloading Excel: {str(e)}")
    
    # Wait for download to complete
    logger.info("Checking for downloaded files...")
    max_wait = 60
    for i in range(max_wait):
        all_files = os.listdir(download_dir)
        downloaded_files = [f for f in all_files if f.endswith('.xls') or f.startswith('vahan')]
        
        if downloaded_files:
            logger.info(f"Download completed successfully: {downloaded_files}")
            break
            
        if i % 10 == 0:  # Log only every 10 seconds to reduce clutter
            logger.info(f"Waiting... {i+1}/{max_wait} seconds")
        time.sleep(1)
    else:
        logger.error("Failed to detect downloaded file after waiting")
        logger.error("Taking screenshot for debugging...")
        try:
            screenshot_path = os.path.join(os.getcwd(), "vahan_error_screenshot.png")
            driver.save_screenshot(screenshot_path)
            logger.info(f"Screenshot saved to {screenshot_path}")
        except Exception as ss_error:
            logger.error(f"Failed to take screenshot: {str(ss_error)}")
        
        # Show page source for debugging
        try:
            with open(os.path.join(os.getcwd(), "page_source.html"), "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            logger.info("Page source saved for debugging")
        except Exception as ps_error:
            logger.error(f"Failed to save page source: {str(ps_error)}")

except Exception as e:
    logger.error(f"An unexpected error occurred: {str(e)}")

finally:
    try:
        logger.info("Closing browser...")
        driver.quit()
    except:
        pass
    logger.info("Script execution completed")