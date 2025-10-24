import pandas as pd
import time
import os
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime

# =============================================================================
# ENHANCED ELIGIBILITY DATA PARSING FUNCTIONS WITH HMO/MA AND MSP EXTRACTION
# =============================================================================

def parse_eligibility_data(file_path):
    """Parse the eligibility Excel file and return structured DataFrame"""
    try:
        df = pd.read_excel(file_path, sheet_name='Eligibility Results')
        
        parsed_records = []
        
        for index, row in df.iterrows():
            if index == 0:  # Skip header
                continue
                
            eligibility_text = str(row.iloc[3])  # Column E
            parsed_data = parse_eligibility_text(eligibility_text)
            
            # Add original patient info
            parsed_data['Original Patient Name'] = row.iloc[0]
            parsed_data['Original Insurance ID'] = row.iloc[1]
            parsed_data['Original Date of Birth'] = row.iloc[2]
            parsed_data['Original Admission Date'] = row.iloc[3]
            parsed_data['Extraction Timestamp'] = row.iloc[-1] if len(row) > 16 else None
            
            parsed_records.append(parsed_data)
        
        return pd.DataFrame(parsed_records)
        
    except Exception as e:
        print(f"Error parsing eligibility data: {e}")
        return pd.DataFrame()

def parse_eligibility_text(text):
    """Parse eligibility response text into structured data"""
    data = {}
    
    # Extract basic information
    patterns = {
        'Beneficiary': r'Beneficiary:\s*([^\n]+)',
        'Sex': r'Sex:\s*([MF])',
        'DOB': r'DOB:\s*([0-9/]+)',
        'Date of Death': r'Date of Death:\s*([^\n]+)',
        'Medicare Number': r'Medicare Number:\s*([^\n]+)',
        'Transaction ID': r'Transaction ID:\s*([^\n]+)',
        'Provider/Supplier': r'Provider/Supplier:\s*([^\n]+)',
        'NPI': r'NPI:\s*([^\n]+)',
        'PTAN': r'PTAN:\s*([^\n]+)',
        'TIN or SSN': r'TIN or SSN:\s*([^\n]+)',
        'From Date of Service': r'From Date of Service:\s*([0-9/]+)',
        'To Date of Service': r'To Date of Service:\s*([0-9/]+)',
        'Part A Effective Date': r'Part A - Beneficiary Details[\s\S]*?Effective Date:\s*([0-9/]+)',
        'Part B Effective Date': r'Part B - Beneficiary Details[\s\S]*?Effective Date:\s*([0-9/]+)',
        'QMB Enrolled': r'QMB Enrolled:\s*([^\n]+)',
        'Base Deductible': r'Base Deductible:\s*([^\n]+)',
        'Remaining Deductible': r'Remaining Deductible:\s*([^\n]+)',
        'Part D Plan Name': r'PBP Plan Name:\s*([^\n]+)'
    }
    
    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        data[key] = match.group(1).strip() if match else None
    
    return data

def extract_hmo_ma_data(driver):
    """Extract data from HMO/MA section"""
    hmo_data = {}
    try:
        print("    Extracting HMO/MA section data...")
        
        # Click on HMO/MA tab using the provided selector
        hmo_tab = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='#hmo']"))
        )
        driver.execute_script("arguments[0].click();", hmo_tab)
        time.sleep(2)  # Wait for content to load
        
        # Extract HMO/MA section content
        hmo_content = driver.find_element(By.ID, "hmo")
        hmo_text = hmo_content.text
        
        # Parse HMO/MA specific data
        hmo_data['HMO_MA_Benefits_Available'] = "No" if "No benefits available" in hmo_text else "Yes"
        
        # Extract additional HMO/MA information
        hmo_patterns = {
            'HMO_MA_Plan_Name': r'Plan Name:\s*([^\n]+)',
            'HMO_MA_Effective_Date': r'Effective Date:\s*([0-9/]+)',
            'HMO_MA_Termination_Date': r'Termination Date:\s*([0-9/]+)',
            'HMO_MA_Plan_ID': r'Plan ID:\s*([^\n]+)',
            'HMO_MA_Group_ID': r'Group ID:\s*([^\n]+)',
            'HMO_MA_Copay_Info': r'Copay:\s*([^\n]+)',
            'HMO_MA_Deductible_Info': r'Deductible:\s*([^\n]+)'
        }
        
        for key, pattern in hmo_patterns.items():
            match = re.search(pattern, hmo_text)
            hmo_data[key] = match.group(1).strip() if match else "Not Available"
        
        print("    ✓ HMO/MA data extracted successfully")
        return hmo_data
        
    except Exception as e:
        print(f"    ✗ Error extracting HMO/MA data: {str(e)}")
        # Return default values if extraction fails
        return {
            'HMO_MA_Benefits_Available': 'Extraction Failed',
            'HMO_MA_Plan_Name': 'Extraction Failed',
            'HMO_MA_Effective_Date': 'Extraction Failed',
            'HMO_MA_Termination_Date': 'Extraction Failed',
            'HMO_MA_Plan_ID': 'Extraction Failed',
            'HMO_MA_Group_ID': 'Extraction Failed',
            'HMO_MA_Copay_Info': 'Extraction Failed',
            'HMO_MA_Deductible_Info': 'Extraction Failed'
        }

def extract_msp_data(driver):
    """Extract data from MSP section"""
    msp_data = {}
    try:
        print("    Extracting MSP section data...")
        
        # Click on MSP tab using the provided selector
        msp_tab = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='#msp']"))
        )
        driver.execute_script("arguments[0].click();", msp_tab)
        time.sleep(2)  # Wait for content to load
        
        # Extract MSP section content
        msp_content = driver.find_element(By.ID, "msp")
        msp_text = msp_content.text
        
        # Parse MSP specific data
        msp_data['MSP_Exists'] = "Yes" if "MSP" in msp_text and "No MSP data" not in msp_text else "No"
        
        # Extract additional MSP information
        msp_patterns = {
            'MSP_Type': r'MSP Type:\s*([^\n]+)',
            'MSP_Effective_Date': r'Effective Date:\s*([0-9/]+)',
            'MSP_Termination_Date': r'Termination Date:\s*([0-9/]+)',
            'MSP_Provider_Name': r'Provider Name:\s*([^\n]+)',
            'MSP_Provider_Phone': r'Provider Phone:\s*([^\n]+)',
            'MSP_Insurance_Name': r'Insurance Name:\s*([^\n]+)',
            'MSP_Insurance_ID': r'Insurance ID:\s*([^\n]+)'
        }
        
        for key, pattern in msp_patterns.items():
            match = re.search(pattern, msp_text)
            msp_data[key] = match.group(1).strip() if match else "Not Available"
        
        print("    ✓ MSP data extracted successfully")
        return msp_data
        
    except Exception as e:
        print(f"    ✗ Error extracting MSP data: {str(e)}")
        # Return default values if extraction fails
        return {
            'MSP_Exists': 'Extraction Failed',
            'MSP_Type': 'Extraction Failed',
            'MSP_Effective_Date': 'Extraction Failed',
            'MSP_Termination_Date': 'Extraction Failed',
            'MSP_Provider_Name': 'Extraction Failed',
            'MSP_Provider_Phone': 'Extraction Failed',
            'MSP_Insurance_Name': 'Extraction Failed',
            'MSP_Insurance_ID': 'Extraction Failed'
        }

def extract_all_tab_data(driver):
    """Extract data from all tabs including HMO/MA and MSP"""
    all_tab_data = {}
    
    try:
        # Extract HMO/MA data
        hmo_data = extract_hmo_ma_data(driver)
        all_tab_data.update(hmo_data)
        
        # Extract MSP data
        msp_data = extract_msp_data(driver)
        all_tab_data.update(msp_data)
        
        # You can add more tab extractions here as needed
        # Example: extract_dsmt_data(driver), extract_mnt_data(driver), etc.
        
    except Exception as e:
        print(f"    ✗ Error extracting tab data: {str(e)}")
    
    return all_tab_data

def clean_eligibility_data(input_file, output_file):
    """Clean and restructure eligibility data after extraction"""
    print("Cleaning and restructuring eligibility data...")
    
    # Parse the data
    parsed_df = parse_eligibility_data(input_file)
    
    if not parsed_df.empty:
        # Save cleaned data
        parsed_df.to_excel(output_file, index=False)
        print(f"✓ Cleaned data saved to: {output_file}")
        print(f"✓ Total records processed: {len(parsed_df)}")
        
        # Print summary
        print("\n=== CLEANED DATA SUMMARY ===")
        for idx, row in parsed_df.iterrows():
            print(f"\nPatient {idx + 1}: {row['Beneficiary']}")
            print(f"  Medicare: {row['Medicare Number']}")
            print(f"  Part A: {row['Part A Effective Date']}")
            print(f"  Part B: {row['Part B Effective Date']}")
            print(f"  QMB: {row['QMB Enrolled']}")
    else:
        print("✗ No data to clean")
    
    return parsed_df

def setup_driver():
    """Setup Chrome driver with minimal options"""
    chrome_options = Options()
    
    # Minimal options to avoid detection and conflicts
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    # Use a simple profile in current directory
    automation_profile = os.path.join(os.getcwd(), "chrome_temp_profile")
    chrome_options.add_argument(f"--user-data-dir={automation_profile}")
    
    try:
        # Use webdriver_manager to automatically handle ChromeDriver
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        
        # Remove webdriver property to avoid detection
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        print("✓ Chrome driver initialized successfully")
        return driver
        
    except Exception as e:
        print(f"✗ Error initializing Chrome driver: {e}")
        raise

def manual_setup_instructions():
    """Display manual setup instructions"""
    print("=" * 70)
    print("ENHANCED MEDICARE ELIGIBILITY AUTOMATION")
    print("=" * 70)
    print("This version now extracts:")
    print("  - Basic eligibility information")
    print("  - HMO/MA section data")
    print("  - MSP (Medicare Secondary Payer) section data")
    print("=" * 70)
    print("STEP 1: VPN CONNECTION")
    print("  - Connect to your organization's VPN manually")
    print("  - Use your regular VPN client/software")
    print("  - Ensure you have stable internet connection through VPN")
    print()
    print("STEP 2: BROWSER PREPARATION")
    print("  - Keep this window open")
    print("  - The script will open Chrome browser automatically")
    print("  - You will manually log into Noridian Medicare Portal")
    print()
    print("STEP 3: LOGIN PROCESS")
    print("  - Manually navigate to: https://www.noridianmedicareportal.com")
    print("  - Enter your credentials and complete login")
    print("  - Stay on the dashboard page")
    print("=" * 70)
    input("Press Enter to continue after you've read the instructions...")

def wait_for_manual_login(driver):
    """Wait for user to manually complete login"""
    print("\n" + "=" * 60)
    print("WAITING FOR MANUAL LOGIN")
    print("=" * 60)
    print("Please complete these steps in the Chrome browser that opened:")
    print("1. If not already there, go to: https://www.noridianmedicareportal.com")
    print("2. Log in with your credentials")
    print("3. Wait until you see the main dashboard")
    print("4. The script will wait 4 minutes before starting automation")
    print("=" * 60)
    
    # Navigate to the portal
    driver.get("https://www.noridianmedicareportal.com")
    
    print("Waiting 4 minutes for you to complete manual login...")
    time.sleep(240)  # Wait 4 minutes for manual login
    
    # Check if we're on a logged-in page
    try:
        # Try multiple indicators of successful login
        logged_in_indicators = [
            "//a[contains(@href, 'eligibility')]",
            "//a[contains(text(), 'Logout')]",
            "//*[contains(text(), 'Welcome')]",
            "//*[contains(text(), 'Dashboard')]",
            "//*[contains(@class, 'user-profile')]",
            "//button[contains(text(), 'Search')]",
            "//*[contains(text(), 'Eligibility')]",
            "//title[contains(text(), 'Portal')]",
            "//*[contains(text(), 'Noridian')]"
        ]
        
        for indicator in logged_in_indicators:
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, indicator))
                )
                print("✓ Login detected! Continuing with automation...")
                return True
            except:
                continue
                
        print("⚠ Could not detect specific login elements, but continuing anyway...")
        return True
        
    except Exception as e:
        print(f"⚠ Continuing despite login detection issues: {str(e)}")
        return True

def detect_eligibility_page_type(driver):
    """Detect what type of eligibility page we're on"""
    try:
        # Check for different form layouts
        
        # Layout 1: Standard form with hicn, lastName, dob fields
        try:
            hicn_field = driver.find_element(By.ID, "hicn")
            last_name_field = driver.find_element(By.ID, "lastName")
            dob_field = driver.find_element(By.ID, "dob")
            print("✓ Detected Standard Eligibility Form Layout")
            return "standard"
        except:
            pass
        
        # Layout 2: MBI Lookup form
        try:
            mbi_field = driver.find_element(By.ID, "mbi")
            print("✓ Detected MBI Lookup Form Layout")
            return "mbi_lookup"
        except:
            pass
        
        # Layout 3: Benefits Inquiry form
        try:
            # Look for elements that might indicate Benefits Inquiry form
            beneficiary_fields = driver.find_elements(By.XPATH, 
                "//input[contains(@name, 'beneficiary') or contains(@id, 'beneficiary') or contains(@placeholder, 'Medicare')]")
            if beneficiary_fields:
                print("✓ Detected Benefits Inquiry Form Layout")
                return "benefits_inquiry"
        except:
            pass
        
        # Layout 4: Check for any form that might be eligibility related
        try:
            form_elements = driver.find_elements(By.XPATH, 
                "//input[@type='text'] | //select | //textarea")
            if len(form_elements) > 2:
                print("✓ Detected Generic Form Layout (will attempt to fill)")
                return "generic_form"
        except:
            pass
        
        print("✗ Could not identify specific eligibility form layout")
        return "unknown"
        
    except Exception as e:
        print(f"✗ Error detecting page type: {str(e)}")
        return "unknown"

def wait_for_manual_eligibility_navigation(driver):
    """Wait for user to manually navigate to eligibility section and detect form type"""
    print("\n" + "=" * 60)
    print("MANUAL ELIGIBILITY NAVIGATION REQUIRED")
    print("=" * 60)
    print("Please manually navigate to the Eligibility section:")
    print("1. Look for 'Eligibility or MBI Lookup' in the menu")
    print("2. Click on 'Eligibility Benefits Inquiry' or similar")
    print("3. Wait for the page to load completely")
    print("4. Make sure you see form fields for patient information")
    print("=" * 60)
    
    # Wait 2 minutes for manual navigation
    print("Waiting 2 minutes for you to navigate to Eligibility page...")
    time.sleep(120)
    
    # Detect what type of page we're on
    page_type = detect_eligibility_page_type(driver)
    
    if page_type != "unknown":
        print(f"✓ Successfully detected {page_type} form layout")
        return True
    else:
        print("✗ Could not detect eligibility form. Please check if you're on the correct page.")
        print("Current page URL:", driver.current_url)
        print("Current page title:", driver.title)
        
        # Ask user to confirm they're on the right page
        confirm = input("Are you on the Eligibility Benefits Inquiry page? (y/n): ").lower().strip()
        if confirm == 'y':
            print("✓ Continuing with automation based on user confirmation")
            return True
        else:
            print("✗ Please navigate to the correct page and run the script again")
            return False

def navigate_to_eligibility(driver):
    """Navigate to Eligibility or MBI Lookup section"""
    max_attempts = 2
    
    for attempt in range(max_attempts):
        try:
            print(f"\nAttempt {attempt + 1} to find Eligibility section...")
            
            # Try multiple selectors for the eligibility link
            selectors = [
                "//a[contains(@href, 'eligibility') and contains(., 'Eligibility')]",
                "//a[contains(text(), 'Eligibility')]",
                "//a[contains(text(), 'MBI Lookup')]",
                "//a[@aria-labelledby='layout_18']",
                "//a[contains(@href, 'eligibility')]",
                "//*[contains(text(), 'Eligibility or MBI Lookup')]",
                "//a[contains(., 'Eligibility')]"
            ]
            
            for selector in selectors:
                try:
                    print(f"  Trying selector: {selector[:50]}...")
                    eligibility_link = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    driver.execute_script("arguments[0].click();", eligibility_link)
                    print("✓ Clicked Eligibility link")
                    
                    # Wait a bit then detect what page we're on
                    time.sleep(5)
                    page_type = detect_eligibility_page_type(driver)
                    
                    if page_type != "unknown":
                        print("✓ Successfully reached eligibility page")
                        return True
                    else:
                        # If we can't detect the form, try to click "Eligibility Benefits Inquiry"
                        try:
                            benefits_link = driver.find_element(By.XPATH, 
                                "//a[contains(text(), 'Eligibility Benefits Inquiry')] | " +
                                "//button[contains(text(), 'Eligibility Benefits')]")
                            driver.execute_script("arguments[0].click();", benefits_link)
                            print("✓ Clicked Eligibility Benefits Inquiry")
                            time.sleep(3)
                            
                            page_type = detect_eligibility_page_type(driver)
                            if page_type != "unknown":
                                return True
                        except:
                            pass
                            
                except Exception as e:
                    continue
                    
            print("✗ Could not find Eligibility link with automated selectors")
            
            # If automated navigation fails, wait for manual navigation
            if attempt == max_attempts - 1:
                print("\nSwitching to manual navigation mode...")
                return wait_for_manual_eligibility_navigation(driver)
            else:
                print("Retrying...")
                time.sleep(5)
                
        except Exception as e:
            print(f"✗ Attempt {attempt + 1} failed: {str(e)}")
            if attempt == max_attempts - 1:
                return wait_for_manual_eligibility_navigation(driver)
    
    return False

def fill_eligibility_form(driver, record):
    """Fill the eligibility form with patient data based on detected form type"""
    try:
        # First detect what type of form we're dealing with
        page_type = detect_eligibility_page_type(driver)
        print(f"  Detected form type: {page_type}")
        
        if page_type == "unknown":
            print("  ✗ Cannot fill form - unknown form layout")
            return False
        
        print(f"  Filling form for: {record['Patient Name']}")
        
        # Extract last name from patient name
        last_name = record['Patient Name'].split(',')[0].strip()
        
        # Convert Date of Birth to proper format
        def convert_date_format(date_str):
            """Convert date from mm/dd/yy to mm/dd/yyyy format"""
            try:
                parts = date_str.split('/')
                if len(parts) == 3:
                    month = parts[0].zfill(2)  # Ensure 2-digit month
                    day = parts[1].zfill(2)    # Ensure 2-digit day
                    year = parts[2]
                    
                    # Handle 2-digit year (convert to 4-digit)
                    if len(year) == 2:
                        # Assuming years 00-25 are 2000-2025, 26-99 are 1926-1999
                        year_int = int(year)
                        if year_int <= 25:
                            year = f"20{year}"
                        else:
                            year = f"19{year}"
                    
                    return f"{month}/{day}/{year}"
                return date_str
            except:
                return date_str
        
        formatted_dob = convert_date_format(record['Date of Birth'])
        print(f"    Original DOB: {record['Date of Birth']}, Formatted: {formatted_dob}")
        
        if page_type == "standard":
            # Standard form with hicn, lastName, dob fields
            hicn_field = driver.find_element(By.ID, "hicn")
            hicn_field.clear()
            hicn_field.send_keys(str(record['Insurance ID']))
            
            last_name_field = driver.find_element(By.ID, "lastName")
            last_name_field.clear()
            last_name_field.send_keys(last_name)
            
            dob_field = driver.find_element(By.ID, "dob")
            dob_field.clear()
            dob_field.send_keys(formatted_dob)  # Use formatted DOB
            
            # Handle date of service selection and fields
            handle_date_of_service(driver, record)
                
        elif page_type in ["benefits_inquiry", "generic_form"]:
            # For benefits inquiry or generic forms, try to find fields by various attributes
            print("    Attempting to fill generic form...")
            
            # Try to find Medicare Number field
            medicare_selectors = [
                "//input[contains(@id, 'hicn')]",
                "//input[contains(@name, 'hicn')]",
                "//input[contains(@placeholder, 'Medicare')]",
                "//input[contains(@id, 'mbi')]",
                "//input[contains(@name, 'mbi')]",
                "//input[@type='text']"  # Fallback to first text input
            ]
            
            for selector in medicare_selectors:
                try:
                    medicare_field = driver.find_element(By.XPATH, selector)
                    medicare_field.clear()
                    medicare_field.send_keys(str(record['Insurance ID']))
                    print(f"    Medicare Number: {record['Insurance ID']}")
                    break
                except:
                    continue
            
            # Try to find Last Name field
            last_name_selectors = [
                "//input[contains(@id, 'lastName')]",
                "//input[contains(@name, 'lastName')]",
                "//input[contains(@placeholder, 'Last Name')]",
                "//input[contains(@id, 'lastname')]",
                "//input[contains(@name, 'lastname')]"
            ]
            
            for selector in last_name_selectors:
                try:
                    last_name_field = driver.find_element(By.XPATH, selector)
                    last_name_field.clear()
                    last_name_field.send_keys(last_name)
                    print(f"    Last Name: {last_name}")
                    break
                except:
                    continue
            
            # Try to find Date of Birth field
            dob_selectors = [
                "//input[contains(@id, 'dob')]",
                "//input[contains(@name, 'dob')]",
                "//input[contains(@placeholder, 'Date of Birth')]",
                "//input[contains(@id, 'birth')]",
                "//input[contains(@name, 'birth')]"
            ]
            
            for selector in dob_selectors:
                try:
                    dob_field = driver.find_element(By.XPATH, selector)
                    dob_field.clear()
                    dob_field.send_keys(formatted_dob)  # Use formatted DOB
                    print(f"    Date of Birth: {formatted_dob}")
                    break
                except:
                    continue
            
            # Handle date of service for generic forms too
            handle_date_of_service(driver, record)
        
        print(f"    ✓ Form filled successfully for {record['Patient Name']}")
        return True
        
    except Exception as e:
        print(f"✗ Error filling form for {record['Patient Name']}: {str(e)}")
        return False

def handle_date_of_service(driver, record):
    """Handle the date of service section - click radio button and fill dates"""
    try:
        print("    Handling date of service section...")
        
        # Convert dates to proper format (mm/dd/yyyy)
        def convert_date_format(date_str):
            """Convert date from mm/dd/yy to mm/dd/yyyy format"""
            try:
                parts = date_str.split('/')
                if len(parts) == 3:
                    month = parts[0].zfill(2)  # Ensure 2-digit month
                    day = parts[1].zfill(2)    # Ensure 2-digit day
                    year = parts[2]
                    
                    # Handle 2-digit year (convert to 4-digit)
                    if len(year) == 2:
                        # Assuming years 00-25 are 2000-2025, 26-99 are 1926-1999
                        year_int = int(year)
                        if year_int <= 25:
                            year = f"20{year}"
                        else:
                            year = f"19{year}"
                    
                    return f"{month}/{day}/{year}"
                return date_str
            except:
                return date_str
        
        # Convert admission date to proper format
        formatted_admission_date = convert_date_format(record['Admission Date'])
        print(f"    Original date: {record['Admission Date']}, Formatted: {formatted_admission_date}")
        
        # Method 1: Click the label that contains the radio button
        label_selectors = [
            "//label[@for='default_date_radio2']",
            "//label[contains(@class, 'radio-icon') and @for='default_date_radio2']",
            "//label[contains(., 'Provide date of service below')]",
            "//label[contains(@class, 'radio-icon') and contains(., 'Provide date of service below')]"
        ]
        
        radio_clicked = False
        for selector in label_selectors:
            try:
                print(f"      Trying label selector: {selector}")
                label_element = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )
                print("      Clicking label element...")
                driver.execute_script("arguments[0].click();", label_element)
                print("    ✓ Selected 'Provide date of service below' via label")
                radio_clicked = True
                time.sleep(2)
                break
            except Exception as e:
                print(f"      Label selector failed: {str(e)}")
                continue
        
        # Method 2: If label clicking doesn't work, try the radio button directly
        if not radio_clicked:
            print("    Trying direct radio button approaches...")
            try:
                radio_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.ID, "default_date_radio2"))
                )
                if not radio_button.is_selected():
                    driver.execute_script("arguments[0].click();", radio_button)
                    print("    ✓ Selected 'Provide date of service below' via radio button")
                    radio_clicked = True
                    time.sleep(2)
            except Exception as e:
                print(f"      Radio button approaches failed: {str(e)}")
        
        if not radio_clicked:
            print("    ⚠ Could not select 'Provide date of service below'")
            return False
        
        # Now fill the fromDate and toDate fields with formatted dates
        print("    Filling date fields...")
        from_date_filled = False
        to_date_filled = False
        
        # Try multiple selectors for fromDate field
        from_date_selectors = [
            "//input[@id='fromDate']",
            "//input[@name='fromDate']",
            "//input[contains(@id, 'fromDate')]",
            "//input[contains(@name, 'fromDate')]",
            "//input[contains(@placeholder, 'From Date')]",
            "//input[contains(@placeholder, 'mm/dd/yyyy')]"
        ]
        
        for selector in from_date_selectors:
            try:
                from_date_field = driver.find_element(By.XPATH, selector)
                from_date_field.clear()
                from_date_field.send_keys(formatted_admission_date)
                print(f"    From Date: {formatted_admission_date}")
                from_date_filled = True
                break
            except:
                continue
        
        # Try multiple selectors for toDate field
        to_date_selectors = [
            "//input[@id='toDate']",
            "//input[@name='toDate']",
            "//input[contains(@id, 'toDate')]",
            "//input[contains(@name, 'toDate')]",
            "//input[contains(@placeholder, 'To Date')]",
            "//input[contains(@placeholder, 'mm/dd/yyyy')]"
        ]
        
        for selector in to_date_selectors:
            try:
                to_date_field = driver.find_element(By.XPATH, selector)
                to_date_field.clear()
                to_date_field.send_keys(formatted_admission_date)
                print(f"    To Date: {formatted_admission_date}")
                to_date_filled = True
                break
            except:
                continue
        
        if not from_date_filled or not to_date_filled:
            print(f"    ⚠ Could not fill date fields - From: {from_date_filled}, To: {to_date_filled}")
            return False
        else:
            print("    ✓ Date of service fields filled successfully")
            return True
            
    except Exception as e:
        print(f"    ⚠ Error handling date of service: {str(e)}")
        return False

def submit_form(driver):
    """Submit the eligibility form"""
    try:
        # Try multiple submit button selectors
        submit_selectors = [
            "//button[@id='btnSubmit']",
            "//input[@type='submit']",
            "//button[contains(text(), 'Submit')]",
            "//button[contains(text(), 'Search')]",
            "//button[contains(text(), 'Inquiry')]",
            "//input[contains(@value, 'Submit')]",
            "//input[contains(@value, 'Search')]"
        ]
        
        for selector in submit_selectors:
            try:
                submit_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )
                driver.execute_script("arguments[0].click();", submit_button)
                print("    ✓ Form submitted")
                
                # Wait for results
                time.sleep(5)
                return True
            except:
                continue
        
        print("✗ Could not find submit button")
        return False
        
    except Exception as e:
        print(f"✗ Error submitting form: {str(e)}")
        return False

def extract_results_data(driver, original_record):
    """Extract data from the results page after submitting inquiry - ENHANCED VERSION"""
    try:
        print("    Extracting results data...")
        
        # Wait for results page to load - look for specific result elements
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//strong[contains(text(), 'Beneficiary:')] | //div[contains(text(), 'Beneficiary:')]"))
        )
        time.sleep(3)
        
        extracted_data = {}
        
        # Extract basic data using existing method
        extracted_data.update(extract_basic_results_data(driver))
        
        # Extract HMO/MA and MSP tab data
        tab_data = extract_all_tab_data(driver)
        extracted_data.update(tab_data)
        
        # Add original record data
        extracted_data['Original Patient Name'] = original_record['Patient Name']
        extracted_data['Original Insurance ID'] = original_record['Insurance ID']
        extracted_data['Original Date of Birth'] = original_record['Date of Birth']
        extracted_data['Original Admission Date'] = original_record['Admission Date']
        extracted_data['Extraction Timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        print("    ✓ All results data extracted successfully")
        return extracted_data
        
    except Exception as e:
        print(f"    ✗ Error extracting results data: {str(e)}")
        return None

def extract_basic_results_data(driver):
    """Extract basic eligibility data (original method)"""
    extracted_data = {}
    
    try:
        # Extract data using more specific selectors based on the actual page structure
        try:
            # Beneficiary Name - try multiple selectors
            beneficiary_selectors = [
                "//strong[contains(text(), 'Beneficiary:')]/following-sibling::text()",
                "//div[contains(., 'Beneficiary:')]",
                "//*[contains(text(), 'Beneficiary:')]",
                "//strong[text()='Beneficiary:']/following-sibling::text()"
            ]
            
            for selector in beneficiary_selectors:
                try:
                    if 'text()' in selector:
                        # For text node selectors
                        beneficiary_text = driver.execute_script(f"""
                            var element = document.evaluate('{selector}', document, null, XPathResult.STRING_TYPE, null).stringValue;
                            return element.trim();
                        """)
                    else:
                        beneficiary_element = driver.find_element(By.XPATH, selector)
                        beneficiary_text = beneficiary_element.text.replace('Beneficiary:', '').strip()
                    
                    if beneficiary_text and len(beneficiary_text) > 1:
                        extracted_data['Beneficiary'] = ' '.join(beneficiary_text.split())
                        break
                except:
                    continue
            
            if 'Beneficiary' not in extracted_data:
                extracted_data['Beneficiary'] = ''
                
        except:
            extracted_data['Beneficiary'] = ''
        
        # Extract other fields using similar approach
        field_mappings = {
            'Sex': ['Sex:', 'Gender:'],
            'DOB': ['DOB:', 'Date of Birth:'],
            'Date of Death': ['Date of Death:'],
            'Medicare Number': ['Medicare Number:'],
            'Transaction ID': ['Transaction ID:'],
            'Provider/Supplier': ['Provider/Supplier:'],
            'NPI': ['NPI:'],
            'PTAN': ['PTAN:'],
            'TIN or SSN': ['TIN or SSN:'],
            'From Date of Service': ['From Date of Service:'],
            'To Date of Service': ['To Date of Service:']
        }
        
        for field_name, possible_labels in field_mappings.items():
            try:
                field_found = False
                for label in possible_labels:
                    try:
                        # Try different XPath patterns to find the field
                        selectors = [
                            f"//strong[contains(text(), '{label}')]/following-sibling::text()[1]",
                            f"//div[contains(., '{label}')]",
                            f"//*[contains(text(), '{label}')]",
                            f"//strong[text()='{label}']/following-sibling::text()"
                        ]
                        
                        for selector in selectors:
                            try:
                                if 'text()' in selector:
                                    # For text node selectors
                                    field_value = driver.execute_script(f"""
                                        var element = document.evaluate('{selector}', document, null, XPathResult.STRING_TYPE, null).stringValue;
                                        return element.trim();
                                    """)
                                else:
                                    field_element = driver.find_element(By.XPATH, selector)
                                    field_value = field_element.text.replace(label, '').strip()
                                
                                if field_value and len(field_value) > 0:
                                    extracted_data[field_name] = field_value
                                    field_found = True
                                    break
                            except:
                                continue
                        
                        if field_found:
                            break
                    except:
                        continue
                
                if not field_found:
                    extracted_data[field_name] = ''
                    
            except:
                extracted_data[field_name] = ''
        
        # Alternative approach: Get all text and parse it
        if not any(extracted_data.values()):
            print("    Trying alternative extraction method...")
            try:
                page_text = driver.find_element(By.TAG_NAME, "body").text
                lines = page_text.split('\n')
                
                for line in lines:
                    for field_name, possible_labels in field_mappings.items():
                        for label in possible_labels:
                            if label in line:
                                value = line.replace(label, '').strip()
                                if value and field_name not in extracted_data:
                                    extracted_data[field_name] = value
                                    break
            except:
                pass
        
        # Debug: Print what was extracted
        print("    Extracted basic data:")
        for key, value in extracted_data.items():
            if value:
                print(f"      {key}: {value}")
                
    except Exception as e:
        print(f"    ✗ Error extracting basic results data: {str(e)}")
    
    return extracted_data

def save_to_excel(extracted_data_list, output_file="Eligibility_Results.xlsx"):
    """Save extracted data to Excel file - ENHANCED VERSION"""
    try:
        if not extracted_data_list:
            print("No data to save")
            return False
        
        # Create DataFrame
        df = pd.DataFrame(extracted_data_list)
        
        # Clean the data - remove any HTML tags or extra whitespace
        for column in df.columns:
            if df[column].dtype == 'object':
                df[column] = df[column].astype(str).str.strip()
        
        # Define the desired column order with new HMO/MA and MSP columns
        column_order = [
            'Original Patient Name', 'Original Insurance ID', 'Original Date of Birth', 'Original Admission Date',
            'Beneficiary', 'Sex', 'DOB', 'Date of Death', 'Medicare Number', 'Transaction ID',
            'Provider/Supplier', 'NPI', 'PTAN', 'TIN or SSN', 'From Date of Service', 'To Date of Service',
            # HMO/MA columns
            'HMO_MA_Benefits_Available', 'HMO_MA_Plan_Name', 'HMO_MA_Effective_Date', 'HMO_MA_Termination_Date',
            'HMO_MA_Plan_ID', 'HMO_MA_Group_ID', 'HMO_MA_Copay_Info', 'HMO_MA_Deductible_Info',
            # MSP columns
            'MSP_Exists', 'MSP_Type', 'MSP_Effective_Date', 'MSP_Termination_Date',
            'MSP_Provider_Name', 'MSP_Provider_Phone', 'MSP_Insurance_Name', 'MSP_Insurance_ID',
            'Extraction Timestamp'
        ]
        
        # Reorder columns and only include existing ones
        existing_columns = [col for col in column_order if col in df.columns]
        remaining_columns = [col for col in df.columns if col not in column_order]
        df = df[existing_columns + remaining_columns]
        
        # Save to Excel with auto-adjust column widths
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Eligibility Results')
            
            # Auto-adjust column widths
            worksheet = writer.sheets['Eligibility Results']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        print(f"✓ Enhanced results saved to {output_file}")
        print(f"✓ Total records saved: {len(extracted_data_list)}")
        print(f"✓ Now includes HMO/MA and MSP section data")
        return True
        
    except Exception as e:
        print(f"✗ Error saving to Excel: {str(e)}")
        return False

def process_excel_data(file_path):
    """Read and process Excel data"""
    try:
        df = pd.read_excel(file_path, sheet_name='Sheet1')
        
        # Remove duplicates based on Patient Name and Insurance ID
        df_unique = df.drop_duplicates(subset=['Patient Name', 'Insurance ID'])
        
        records = df_unique.to_dict('records')
        print(f"✓ Found {len(records)} unique patients to process")
        return records
    except Exception as e:
        print(f"✗ Error reading Excel file: {str(e)}")
        return []

# =============================================================================
# MAIN FUNCTION
# =============================================================================

def main():
    # Configuration
    EXCEL_FILE_PATH = "Input_Details.xlsx"
    OUTPUT_FILE = "Eligibility_Results.xlsx"
    CLEANED_OUTPUT_FILE = "Cleaned_Eligibility_Results.xlsx"
    
    # Check if Excel file exists
    if not os.path.exists(EXCEL_FILE_PATH):
        print(f"✗ Excel file not found: {EXCEL_FILE_PATH}")
        print("Please make sure the Excel file is in the same directory as the script")
        return
    
    # Ask user what they want to do
    print("=" * 70)
    print("ENHANCED MEDICARE ELIGIBILITY AUTOMATION TOOL")
    print("=" * 70)
    print("Now extracts: Basic Info + HMO/MA + MSP Section Data")
    print("=" * 70)
    print("Choose an option:")
    print("1. Run full automation (extract data from Noridian portal)")
    print("2. Clean existing eligibility data (parse already extracted data)")
    print("=" * 70)
    
    choice = input("Enter your choice (1 or 2): ").strip()
    
    if choice == "2":
        # Clean existing eligibility data
        if os.path.exists(OUTPUT_FILE):
            clean_eligibility_data(OUTPUT_FILE, CLEANED_OUTPUT_FILE)
        else:
            print(f"✗ Eligibility results file not found: {OUTPUT_FILE}")
            print("Please run option 1 first to extract data from the portal")
        return
    
    elif choice != "1":
        print("✗ Invalid choice. Exiting.")
        return
    
    # Process Excel data for full automation
    records = process_excel_data(EXCEL_FILE_PATH)
    if not records:
        print("✗ No records to process")
        return
    
    # Display manual setup instructions
    manual_setup_instructions()
    
    driver = None
    extracted_data_list = []
    
    try:
        # Setup driver
        print("\nInitializing Chrome browser...")
        driver = setup_driver()
        
        # Wait for manual login
        if not wait_for_manual_login(driver):
            print("⚠ Continuing despite login detection issues...")
        
        # Add additional wait time before starting automation
        print("\n" + "=" * 50)
        print("STARTING AUTOMATION IN 30 SECONDS...")
        print("=" * 50)
        time.sleep(30)
        
        # Process each record
        successful = 0
        failed = 0
        
        for i, record in enumerate(records, 1):
            print(f"\n{'='*50}")
            print(f"Processing patient {i}/{len(records)}: {record['Patient Name']}")
            print(f"{'='*50}")
            
            # Navigate to eligibility section
            if not navigate_to_eligibility(driver):
                print("✗ Failed to navigate to eligibility section")
                failed += 1
                continue
            
            # Fill and submit form
            if fill_eligibility_form(driver, record) and submit_form(driver):
                # Extract results data (now includes HMO/MA and MSP)
                extracted_data = extract_results_data(driver, record)
                if extracted_data:
                    extracted_data_list.append(extracted_data)
                    successful += 1
                    print(f"✓ Successfully processed and extracted: {record['Patient Name']}")
                    
                    # Save progress after each successful extraction
                    save_to_excel(extracted_data_list, OUTPUT_FILE)
                else:
                    failed += 1
                    print(f"✗ Processed but failed to extract: {record['Patient Name']}")
            else:
                failed += 1
                print(f"✗ Failed to process: {record['Patient Name']}")
            
            # Wait before next patient
            time.sleep(3)
            
        # Print summary
        print(f"\n{'='*60}")
        print("PROCESSING SUMMARY")
        print(f"{'='*60}")
        print(f"Total patients: {len(records)}")
        print(f"Successful: {successful}")
        print(f"Failed: {failed}")
        print(f"Enhanced data saved to: {OUTPUT_FILE}")
        print(f"Now includes: Basic Info + HMO/MA + MSP Section Data")
        print(f"{'='*60}")
        
        # Offer to clean the data after extraction
        if successful > 0:
            clean_choice = input("\nDo you want to clean and restructure the extracted data? (y/n): ").lower().strip()
            if clean_choice == 'y':
                clean_eligibility_data(OUTPUT_FILE, CLEANED_OUTPUT_FILE)
            
    except Exception as e:
        print(f"\n✗ An error occurred: {str(e)}")
    finally:
        if driver:
            print("\nProcessing completed.")
            keep_open = input("Keep browser open? (y/n): ").lower().strip()
            if keep_open != 'y':
                driver.quit()
                print("Browser closed.")
            else:
                print("Browser remains open. You can close it manually when done.")

if __name__ == "__main__":
    main()