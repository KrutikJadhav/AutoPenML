import pandas as pd
import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime

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

def extract_results_data(driver, original_record):
    """Extract data from the results page after submitting inquiry"""
    try:
        print("    Extracting results data...")
        
        # Wait for results page to load
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'col-4')]//strong"))
        )
        time.sleep(3)
        
        extracted_data = {}
        
        # Extract data from first column (Beneficiary information)
        try:
            # Beneficiary Name
            beneficiary_element = driver.find_element(By.XPATH, "//div[contains(., 'Beneficiary:')]")
            beneficiary_text = beneficiary_element.text.replace('Beneficiary:', '').strip()
            extracted_data['Beneficiary'] = ' '.join(beneficiary_text.split())
        except:
            extracted_data['Beneficiary'] = ''
        
        try:
            # Sex
            sex_element = driver.find_element(By.XPATH, "//div[contains(., 'Sex:')]")
            extracted_data['Sex'] = sex_element.text.replace('Sex:', '').strip()
        except:
            extracted_data['Sex'] = ''
        
        try:
            # DOB
            dob_element = driver.find_element(By.XPATH, "//div[contains(., 'DOB:')]")
            extracted_data['DOB'] = dob_element.text.replace('DOB:', '').strip()
        except:
            extracted_data['DOB'] = ''
        
        try:
            # Date of Death
            death_element = driver.find_element(By.XPATH, "//div[contains(., 'Date of Death:')]")
            extracted_data['Date of Death'] = death_element.text.replace('Date of Death:', '').strip()
        except:
            extracted_data['Date of Death'] = ''
        
        try:
            # Medicare Number
            medicare_element = driver.find_element(By.XPATH, "//div[contains(., 'Medicare Number:')]")
            extracted_data['Medicare Number'] = medicare_element.text.replace('Medicare Number:', '').strip()
        except:
            extracted_data['Medicare Number'] = ''
        
        try:
            # Transaction ID
            transaction_element = driver.find_element(By.XPATH, "//div[contains(., 'Transaction ID:')]")
            extracted_data['Transaction ID'] = transaction_element.text.replace('Transaction ID:', '').strip()
        except:
            extracted_data['Transaction ID'] = ''
        
        # Extract data from second column (Provider information)
        try:
            # Provider/Supplier
            provider_element = driver.find_element(By.XPATH, "//div[contains(., 'Provider/Supplier:')]")
            extracted_data['Provider/Supplier'] = provider_element.text.replace('Provider/Supplier:', '').strip()
        except:
            extracted_data['Provider/Supplier'] = ''
        
        try:
            # NPI
            npi_element = driver.find_element(By.XPATH, "//div[contains(., 'NPI:')]")
            extracted_data['NPI'] = npi_element.text.replace('NPI:', '').strip()
        except:
            extracted_data['NPI'] = ''
        
        try:
            # PTAN
            ptan_element = driver.find_element(By.XPATH, "//div[contains(., 'PTAN:')]")
            extracted_data['PTAN'] = ptan_element.text.replace('PTAN:', '').strip()
        except:
            extracted_data['PTAN'] = ''
        
        try:
            # TIN or SSN
            tin_element = driver.find_element(By.XPATH, "//div[contains(., 'TIN or SSN:')]")
            extracted_data['TIN or SSN'] = tin_element.text.replace('TIN or SSN:', '').strip()
        except:
            extracted_data['TIN or SSN'] = ''
        
        try:
            # From Date of Service
            from_date_element = driver.find_element(By.XPATH, "//div[contains(., 'From Date of Service:')]")
            extracted_data['From Date of Service'] = from_date_element.text.replace('From Date of Service:', '').strip()
        except:
            extracted_data['From Date of Service'] = ''
        
        try:
            # To Date of Service
            to_date_element = driver.find_element(By.XPATH, "//div[contains(., 'To Date of Service:')]")
            extracted_data['To Date of Service'] = to_date_element.text.replace('To Date of Service:', '').strip()
        except:
            extracted_data['To Date of Service'] = ''
        
        # Add original record data
        extracted_data['Original Patient Name'] = original_record['Patient Name']
        extracted_data['Original Insurance ID'] = original_record['Insurance ID']
        extracted_data['Original Date of Birth'] = original_record['Date of Birth']
        extracted_data['Original Admission Date'] = original_record['Admission Date']
        extracted_data['Extraction Timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        print("    ✓ Results data extracted successfully")
        return extracted_data
        
    except Exception as e:
        print(f"    ✗ Error extracting results data: {str(e)}")
        return None

def save_to_excel(extracted_data_list, output_file="eligibility_results.xlsx"):
    """Save extracted data to Excel file"""
    try:
        if not extracted_data_list:
            print("No data to save")
            return False
        
        # Create DataFrame
        df = pd.DataFrame(extracted_data_list)
        
        # Reorder columns to make it more readable
        column_order = [
            'Original Patient Name', 'Original Insurance ID', 'Original Date of Birth', 'Original Admission Date',
            'Beneficiary', 'Sex', 'DOB', 'Date of Death', 'Medicare Number', 'Transaction ID',
            'Provider/Supplier', 'NPI', 'PTAN', 'TIN or SSN', 'From Date of Service', 'To Date of Service',
            'Extraction Timestamp'
        ]
        
        # Only include columns that exist in the DataFrame
        existing_columns = [col for col in column_order if col in df.columns]
        df = df[existing_columns + [col for col in df.columns if col not in column_order]]
        
        # Save to Excel
        df.to_excel(output_file, index=False)
        print(f"✓ Results saved to {output_file}")
        print(f"✓ Total records saved: {len(extracted_data_list)}")
        return True
        
    except Exception as e:
        print(f"✗ Error saving to Excel: {str(e)}")
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

# ... (keep all the other existing functions: setup_driver, manual_setup_instructions, wait_for_manual_login, detect_eligibility_page_type, wait_for_manual_eligibility_navigation, navigate_to_eligibility, fill_eligibility_form, submit_form, process_excel_data)

def main():
    # Configuration
    EXCEL_FILE_PATH = "Input_Details.xlsx"
    OUTPUT_FILE = "Eligibility_Results.xlsx"
    
    # Check if Excel file exists
    if not os.path.exists(EXCEL_FILE_PATH):
        print(f"✗ Excel file not found: {EXCEL_FILE_PATH}")
        print("Please make sure the Excel file is in the same directory as the script")
        return
    
    # Process Excel data
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
                # Extract results data
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
        print(f"Data saved to: {OUTPUT_FILE}")
        print(f"{'='*60}")
            
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