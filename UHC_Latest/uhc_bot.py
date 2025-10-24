import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import openpyxl
from datetime import datetime
import os

class UnitedHealthcareBot:
    def __init__(self):
        self.driver = None
        self.wait = None
        self.setup_driver()
    
    def setup_driver(self):
        """Setup Chrome driver with appropriate options"""
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_argument('--start-maximized')
        
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        self.wait = WebDriverWait(self.driver, 15)
        print("Chrome driver setup completed successfully!")
    
    def login(self, username, password, login_url):
        """Login to United Healthcare portal"""
        try:
            print("Navigating to United Healthcare portal...")
            self.driver.get(login_url)
            
            # Wait for login page to load
            time.sleep(5)
            
            print("Please login manually...")
            input("After logging in successfully and reaching the dashboard, press Enter to continue...")
            
            return True
        except Exception as e:
            print(f"Error during login navigation: {e}")
            return False
    
    def click_eligibility_section(self):
        """Click the eligibility section to load the form"""
        try:
            print("Clicking Eligibility section...")
            
            # Try different possible ways to find and click eligibility section
            eligibility_selectors = [
                "//a[contains(., 'Eligibility')]",
                "//span[contains(., 'Eligibility')]",
                "//button[contains(., 'Eligibility')]",
                "//*[@data-testid='eligibility-tab']",
                "//*[contains(@class, 'eligibility')]",
                "//a[contains(@href, 'eligibility')]",
                "//*[contains(text(), 'Eligibility & Benefits')]",
                "//*[contains(text(), 'Eligibility') and contains(text(), 'Benefits')]",
                "//li[contains(., 'Eligibility')]",
                "//div[contains(., 'Eligibility')]"
            ]
            
            for selector in eligibility_selectors:
                try:
                    print(f"Trying selector: {selector}")
                    eligibility_element = self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    eligibility_element.click()
                    print(f"✓ Successfully clicked Eligibility section using: {selector}")
                    time.sleep(3)
                    return True
                except Exception as e:
                    print(f"✗ Failed with {selector}: {e}")
                    continue
            
            print("❌ Could not find Eligibility section with any selector")
            return False
            
        except Exception as e:
            print(f"Error clicking eligibility section: {e}")
            return False
    
    def navigate_to_eligibility(self):
        """Navigate to eligibility verification section (first time)"""
        return self.click_eligibility_section()
    
    def navigate_back_to_eligibility(self):
        """Navigate back to eligibility section after processing a patient"""
        try:
            print("Navigating back to eligibility section for next patient...")
            
            # First, check if we're already on a page with the eligibility form visible
            try:
                # Check if the member ID field is already visible
                member_id_field = self.driver.find_element(By.ID, "eligibility-memberid-input")
                if member_id_field.is_displayed():
                    print("✓ Eligibility form is already visible")
                    return True
            except:
                pass
            
            # If form is not visible, we need to click the Eligibility section again
            if not self.click_eligibility_section():
                print("❌ Could not navigate back to eligibility section")
                return False
            
            # Wait for the form to load
            time.sleep(3)
            
            # Verify the form is loaded by checking for the member ID field
            try:
                member_id_field = self.wait.until(
                    EC.presence_of_element_located((By.ID, "eligibility-memberid-input"))
                )
                print("✓ Eligibility form loaded successfully")
                return True
            except:
                print("❌ Eligibility form did not load after clicking section")
                return False
            
        except Exception as e:
            print(f"Error navigating back to eligibility: {e}")
            return False
    
    def clear_field_advanced(self, field_element, field_name):
        """Clear a field using multiple advanced methods"""
        try:
            print(f"Clearing {field_name} field...")
            
            # Method 1: JavaScript clearing (most reliable)
            try:
                self.driver.execute_script("arguments[0].value = '';", field_element)
                print(f"✓ {field_name} cleared with JavaScript")
            except Exception as e:
                print(f"JavaScript clearing failed: {e}")
            
            # Method 2: Select all and delete
            try:
                field_element.click()
                time.sleep(0.5)
                field_element.send_keys(Keys.CONTROL + "a")
                time.sleep(0.5)
                field_element.send_keys(Keys.DELETE)
                print(f"✓ {field_name} cleared with select-all + delete")
            except Exception as e:
                print(f"Select-all clearing failed: {e}")
            
            # Method 3: Multiple backspaces
            try:
                current_value = field_element.get_attribute('value')
                if current_value:
                    field_element.click()
                    time.sleep(0.5)
                    for _ in range(len(current_value) + 5):  # Extra backspaces for safety
                        field_element.send_keys(Keys.BACKSPACE)
                        time.sleep(0.05)
                    print(f"✓ {field_name} cleared with backspaces")
            except Exception as e:
                print(f"Backspace clearing failed: {e}")
            
            # Method 4: Clear using Actions chain
            try:
                actions = ActionChains(self.driver)
                actions.click(field_element)
                actions.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL)
                actions.send_keys(Keys.DELETE)
                actions.perform()
                print(f"✓ {field_name} cleared with Actions chain")
            except Exception as e:
                print(f"Actions chain clearing failed: {e}")
            
            # Verify the field is actually empty
            final_value = field_element.get_attribute('value')
            if final_value and final_value != '__/__/____':  # Special case for DOB placeholder
                print(f"⚠️ {field_name} still has value after clearing: '{final_value}'")
                # Try one more time with JavaScript
                self.driver.execute_script("arguments[0].value = '';", field_element)
            else:
                print(f"✓ {field_name} verified empty")
            
            time.sleep(0.5)
            return True
            
        except Exception as e:
            print(f"Error clearing {field_name}: {e}")
            return False
    
    def fill_field_advanced(self, field_element, value, field_name):
        """Fill a field using multiple methods"""
        try:
            print(f"Filling {field_name} with: {value}")
            
            # Method 1: Direct send_keys
            try:
                field_element.click()
                time.sleep(0.5)
                field_element.send_keys(value)
                print(f"✓ {field_name} filled with direct send_keys")
            except Exception as e:
                print(f"Direct send_keys failed: {e}")
            
            # Method 2: JavaScript injection
            try:
                self.driver.execute_script("arguments[0].value = arguments[1];", field_element, value)
                print(f"✓ {field_name} filled with JavaScript")
            except Exception as e:
                print(f"JavaScript filling failed: {e}")
            
            # Method 3: Character by character with delays (for problematic fields)
            try:
                current_value = field_element.get_attribute('value')
                if current_value != value:
                    field_element.click()
                    time.sleep(0.5)
                    field_element.clear()
                    time.sleep(0.5)
                    for char in value:
                        field_element.send_keys(char)
                        time.sleep(0.1)
                    print(f"✓ {field_name} filled character by character")
            except Exception as e:
                print(f"Character-by-character filling failed: {e}")
            
            # Verify the field has the correct value
            final_value = field_element.get_attribute('value')
            if final_value == value:
                print(f"✓ {field_name} verified: '{final_value}'")
                return True
            else:
                print(f"⚠️ {field_name} value mismatch. Expected: '{value}', Got: '{final_value}'")
                return False
                
        except Exception as e:
            print(f"Error filling {field_name}: {e}")
            return False
    
    def find_and_clear_form(self):
        """Find and completely clear the form"""
        try:
            print("Finding and clearing form fields...")
            
            # Find Member ID field with multiple selectors
            member_id_selectors = [
                "//input[@id='eligibility-memberid-input']",
                "//input[contains(@id, 'memberid')]",
                "//input[contains(@id, 'member-id')]",
                "//input[contains(@name, 'memberid')]",
                "//input[@placeholder*='Member ID']",
                "//input[@data-testid*='member-id']",
                "//input[contains(@data-testid, 'eligibility-search-member-id')]"
            ]
            
            member_id_field = None
            for selector in member_id_selectors:
                try:
                    member_id_field = self.driver.find_element(By.XPATH, selector)
                    print(f"✓ Found Member ID field with: {selector}")
                    break
                except:
                    continue
            
            if not member_id_field:
                print("❌ Could not find Member ID field with any selector")
                return False
            
            # Find Date of Birth field with multiple selectors
            dob_selectors = [
                "//input[@id='eligibility-dateofbirth-input']",
                "//input[contains(@id, 'dateofbirth')]",
                "//input[contains(@id, 'dob')]",
                "//input[contains(@name, 'dateofbirth')]",
                "//input[contains(@name, 'dob')]",
                "//input[@placeholder*='Date of Birth']",
                "//input[@placeholder*='DOB']",
                "//input[@data-testid*='dob']",
                "//input[contains(@data-testid, 'eligibility-search-DOB')]",
                "//input[@type='date']"
            ]
            
            dob_field = None
            for selector in dob_selectors:
                try:
                    dob_field = self.driver.find_element(By.XPATH, selector)
                    print(f"✓ Found Date of Birth field with: {selector}")
                    break
                except:
                    continue
            
            if not dob_field:
                print("❌ Could not find Date of Birth field with any selector")
                return False
            
            # Clear both fields using advanced methods
            self.clear_field_advanced(member_id_field, "Member ID")
            self.clear_field_advanced(dob_field, "Date of Birth")
            
            return member_id_field, dob_field
            
        except Exception as e:
            print(f"Error finding and clearing form: {e}")
            return False
    
    def select_custom_date_radio(self):
        """Select the custom date radio button"""
        try:
            print("Selecting custom date radio button...")
            
            # Try multiple selectors for custom radio button
            radio_selectors = [
                "//input[@value='custom']",
                "//input[@name='custom']",
                "//input[contains(@id, 'custom')]",
                "//input[contains(@name, 'custom')]",
                "//input[@type='radio'][2]",  # Second radio button
                "//input[@type='radio'][last()]",  # Last radio button
                "//label[contains(., 'Custom Date')]",
                "//*[contains(text(), 'Custom Date')]//preceding-sibling::input",
                "//*[contains(text(), 'Custom Date')]",
            ]
            
            for selector in radio_selectors:
                try:
                    radio_element = self.driver.find_element(By.XPATH, selector)
                    
                    # If it's a label, click it
                    if radio_element.tag_name.lower() == 'label':
                        radio_element.click()
                        print(f"✓ Custom date selected via label: {selector}")
                        return True
                    else:
                        # If it's an input, check if it's already selected
                        if not radio_element.is_selected():
                            try:
                                radio_element.click()
                                print(f"✓ Custom date selected via input: {selector}")
                                return True
                            except:
                                # Try clicking via JavaScript
                                self.driver.execute_script("arguments[0].click();", radio_element)
                                print(f"✓ Custom date selected via JavaScript: {selector}")
                                return True
                        else:
                            print(f"✓ Custom date already selected: {selector}")
                            return True
                except:
                    continue
            
            print("⚠️ Could not find custom date radio button, but continuing...")
            return True  # Continue even if custom radio not found
            
        except Exception as e:
            print(f"Error selecting custom date: {e}")
            return True  # Continue even if there's an error
    
    def enter_member_info(self, member_id, date_of_birth):
        """Enter member ID and date of birth using advanced methods"""
        try:
            print(f"Entering member info - ID: {member_id}, DOB: {date_of_birth}")
            
            # Find and clear the form completely
            form_fields = self.find_and_clear_form()
            if not form_fields:
                return False
            
            member_id_field, dob_field = form_fields
            
            # Fill Member ID field
            if not self.fill_field_advanced(member_id_field, member_id, "Member ID"):
                print("❌ Failed to fill Member ID")
                return False
            
            # Fill Date of Birth field
            if not self.fill_field_advanced(dob_field, date_of_birth, "Date of Birth"):
                print("❌ Failed to fill Date of Birth")
                return False
            
            # Select Custom Date radio button
            self.select_custom_date_radio()
            
            # Final verification
            time.sleep(1)
            final_member_id = member_id_field.get_attribute('value')
            final_dob = dob_field.get_attribute('value')
            
            print(f"Final verification - Member ID: '{final_member_id}', DOB: '{final_dob}'")
            
            if final_member_id != member_id:
                print(f"❌ Member ID final mismatch: Expected '{member_id}', Got '{final_member_id}'")
                return False
            
            if final_dob != date_of_birth:
                print(f"❌ DOB final mismatch: Expected '{date_of_birth}', Got '{final_dob}'")
                return False
            
            print("✓ All fields filled and verified successfully")
            return True
            
        except Exception as e:
            print(f"Error entering member info: {e}")
            return False
    
    def verify_eligibility(self):
        """Click verify eligibility button using multiple methods"""
        try:
            print("Looking for verify eligibility button...")
            
            # Wait for button to be clickable
            time.sleep(2)
            
            # Try multiple selectors for verify button
            button_selectors = [
                "//button[@id='submit-search-button']",
                "//button[contains(., 'Verify Eligibility')]",
                "//*[contains(text(), 'Verify Eligibility')]",
                "//input[@value='Verify Eligibility']",
                "//button[contains(@class, 'abyss-button-root')]",
                "//button[contains(@type, 'submit')]",
                "//button[contains(@type, 'button') and contains(., 'Verify')]"
            ]
            
            for selector in button_selectors:
                try:
                    verify_button = self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    
                    # Check if button is enabled
                    if verify_button.get_attribute("aria-disabled") == "true":
                        print(f"✗ Verify button is disabled with selector: {selector}")
                        continue
                    
                    # Scroll into view and click
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", verify_button)
                    time.sleep(1)
                    
                    # Try multiple click methods
                    try:
                        verify_button.click()
                        print(f"✓ Verify button clicked with selector: {selector}")
                    except:
                        self.driver.execute_script("arguments[0].click();", verify_button)
                        print(f"✓ Verify button clicked with JavaScript: {selector}")
                    
                    # Wait for results to load
                    print("Waiting for results to load...")
                    time.sleep(5)
                    
                    return True
                    
                except TimeoutException:
                    continue
                except Exception as e:
                    print(f"Error with selector {selector}: {e}")
                    continue
            
            print("✗ Verify button not found with any method")
            return False
            
        except Exception as e:
            print(f"Error clicking verify button: {e}")
            return False
    
    def extract_eligibility_details(self):
        """Extract eligibility details using specific selectors"""
        try:
            print("Extracting eligibility details...")
            
            # Wait a bit for results to load
            time.sleep(3)
            
            # Check for "no results" or error messages first
            no_result_indicators = [
                "//*[contains(., 'not found')]",
                "//*[contains(., 'no results')]",
                "//*[contains(., 'invalid')]",
                "//*[contains(., 'error')]",
                "//*[contains(., 'no member')]",
                "//*[contains(., 'no records')]",
                "//*[contains(@class, 'error')]",
                "//*[contains(@class, 'no-results')]"
            ]
            
            for selector in no_result_indicators:
                try:
                    error_element = self.driver.find_element(By.XPATH, selector)
                    if error_element.is_displayed():
                        error_text = error_element.text.lower()
                        if any(term in error_text for term in ['not found', 'no results', 'invalid', 'no member']):
                            print("✗ No member found or invalid data")
                            return "No User Found"
                except:
                    continue
            
            # Initialize details dictionary with all required fields
            details = {
                'plan_name': 'Not Found',
                'funding_type': 'Not Found',
                'group': 'Not Found',
                'plan_type': 'Not Found',
                'payer_status': 'Not Found',
                'is_active': 'Not Found',
                'extraction_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
            # Extract Plan Name
            try:
                plan_name_element = self.driver.find_element(By.XPATH, "//span[@class='abyss-c-cQFdVt' and contains(text(), 'Choice') or contains(text(), 'Plus') or contains(text(), 'Plan')]")
                details['plan_name'] = plan_name_element.text
                print(f"✓ Plan Name: {details['plan_name']}")
            except:
                print("✗ Plan Name not found")
            
            # Extract Funding Type
            try:
                funding_elements = self.driver.find_elements(By.XPATH, "//span[@class='abyss-c-cQFdVt']")
                for element in funding_elements:
                    text = element.text
                    if 'Self Insured' in text or 'Large Group' in text or 'Funded' in text:
                        details['funding_type'] = text
                        print(f"✓ Funding Type: {details['funding_type']}")
                        break
            except:
                print("✗ Funding Type not found")
            
            # Extract Group Number
            try:
                group_element = self.driver.find_element(By.XPATH, "//p[@class='abyss-c-cQFdVt']")
                details['group'] = group_element.text
                print(f"✓ Group: {details['group']}")
            except:
                print("✗ Group not found")
            
            # Extract Plan Type
            try:
                plan_type_elements = self.driver.find_elements(By.XPATH, "//span[@class='abyss-c-cQFdVt']")
                for element in plan_type_elements:
                    text = element.text
                    if 'Commercial' in text or 'HMO' in text or 'PPO' in text or 'EPO' in text:
                        details['plan_type'] = text
                        print(f"✓ Plan Type: {details['plan_type']}")
                        break
            except:
                print("✗ Plan Type not found")
            
            # Extract Payer Status
            try:
                payer_status_element = self.driver.find_element(By.XPATH, "//strong[contains(text(), 'Primary') or contains(text(), 'Secondary') or contains(text(), 'Tertiary')]")
                details['payer_status'] = payer_status_element.text
                print(f"✓ Payer Status: {details['payer_status']}")
            except:
                print("✗ Payer Status not found")
            
            # Determine if active based on the presence of data
            if (details['plan_name'] != 'Not Found' or 
                details['funding_type'] != 'Not Found' or 
                details['group'] != 'Not Found'):
                details['is_active'] = 'Yes'
            else:
                details['is_active'] = 'No'
            
            print(f"✓ Active Status: {details['is_active']}")
            
            # Print summary of extracted data
            print("\n📊 EXTRACTION SUMMARY:")
            print(f"   Plan Name: {details['plan_name']}")
            print(f"   Funding Type: {details['funding_type']}")
            print(f"   Group: {details['group']}")
            print(f"   Plan Type: {details['plan_type']}")
            print(f"   Payer Status: {details['payer_status']}")
            print(f"   Is Active: {details['is_active']}")
            
            return details
            
        except Exception as e:
            print(f"Error extracting details: {e}")
            return "Error during extraction"
    
    def read_excel_data(self, file_path):
        """Read patient data from Excel file"""
        try:
            # Try to read with specific sheet name first, then fallback to first sheet
            try:
                df = pd.read_excel(file_path, sheet_name='UHC Members')
                print("✓ Read data from 'UHC Members' sheet")
            except:
                df = pd.read_excel(file_path, sheet_name=0)  # First sheet
                print("✓ Read data from first sheet")
            
            print(f"Successfully read Excel file with {len(df)} patients")
            
            # Check if required columns exist
            if 'Member_ID' not in df.columns:
                print("❌ Error: 'Member_ID' column not found in Excel file")
                return None
            if 'Date_of_Birth' not in df.columns:
                print("❌ Error: 'Date_of_Birth' column not found in Excel file")
                return None
            
            return df
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return None
    
    def process_patient(self, member_id, date_of_birth):
        """Process a single patient's eligibility verification"""
        print(f"\n=== Processing patient: {member_id} ===")
        
        # Enter member information
        if not self.enter_member_info(member_id, date_of_birth):
            return "Error: Could not enter member info"
        
        # Verify eligibility
        if not self.verify_eligibility():
            return "Error: Could not verify eligibility"
        
        # Extract details
        details = self.extract_eligibility_details()
        
        # Navigate back to eligibility section for next patient
        if not self.navigate_back_to_eligibility():
            return "Error: Could not navigate back to eligibility"
        
        return details
    
    def write_results_to_excel(self, input_file, output_file, results):
        """Write results back to Excel file"""
        try:
            # Read original data
            df = pd.read_excel(input_file, sheet_name='UHC Members')
            
            # Add result columns if they don't exist
            result_columns = [
                'plan_name', 'funding_type', 'group', 'plan_type', 
                'payer_status', 'is_active', 'processing_status', 'extraction_time'
            ]
            
            for col in result_columns:
                if col not in df.columns:
                    df[col] = ""
            
            # Update with results
            for i, result in enumerate(results):
                if i < len(df):
                    if isinstance(result, dict):
                        for key, value in result.items():
                            if key in df.columns:
                                df.at[i, key] = value
                        df.at[i, 'processing_status'] = 'Completed'
                    else:
                        df.at[i, 'processing_status'] = result
                        for col in ['plan_name', 'funding_type', 'group', 'plan_type', 'payer_status', 'is_active']:
                            if col in df.columns:
                                df.at[i, col] = 'N/A'
            
            # Save to output file
            df.to_excel(output_file, index=False)
            print(f"✓ Results saved to: {output_file}")
            
        except Exception as e:
            print(f"Error writing to Excel: {e}")
    
    def run_automation(self, excel_file_path, output_file_path, login_url):
        """Main automation function"""
        # Read patient data
        patient_data = self.read_excel_data(excel_file_path)
        if patient_data is None:
            return
        
        # Login to portal
        if not self.login("", "", login_url):
            print("Login failed. Exiting...")
            return
        
        # Navigate to eligibility section (first time)
        if not self.navigate_to_eligibility():
            print("Failed to navigate to eligibility section. Exiting...")
            return
        
        results = []
        total_patients = len(patient_data)
        
        print(f"\nStarting automation for {total_patients} patients...")
        
        for index, row in patient_data.iterrows():
            # Get member data using exact column names
            member_id = str(row['Member_ID']).strip()
            date_of_birth = str(row['Date_of_Birth']).strip()
            
            # Handle NaN values and empty strings
            if pd.isna(row['Member_ID']) or member_id == 'nan' or member_id == 'None' or not member_id:
                print(f"Skipping row {index + 1} - missing member ID")
                results.append("Error: Missing Member ID")
                continue
            
            if pd.isna(row['Date_of_Birth']) or date_of_birth == 'nan' or date_of_birth == 'None' or not date_of_birth:
                print(f"Skipping row {index + 1} - missing date of birth")
                results.append("Error: Missing Date of Birth")
                continue
            
            print(f"\n[{index + 1}/{total_patients}] Processing Member ID: {member_id}, DOB: {date_of_birth}")
            
            # Process patient
            result = self.process_patient(member_id, date_of_birth)
            results.append(result)
            
            # Add delay between patients
            time.sleep(3)
            
            print(f"Result: {result}")
        
        # Write results to Excel
        self.write_results_to_excel(excel_file_path, output_file_path, results)
        
        print("\n" + "="*50)
        print("🎉 AUTOMATION COMPLETED SUCCESSFULLY!")
        print("="*50)
    
    def close(self):
        """Close the browser"""
        if self.driver:
            self.driver.quit()
            print("Browser closed.")

# Main execution
if __name__ == "__main__":
    print("Starting United Healthcare Automation Bot...")
    print("=" * 50)
    
    bot = UnitedHealthcareBot()
    
    try:
        # Configuration
        input_excel = "patient_list.xlsx"  # Your Excel file name
        output_excel = "eligibility_results.xlsx"  # Output file
        uhc_login_url = "https://www.uhcprovider.com"  # UHC portal URL
        
        print(f"Input file: {input_excel}")
        print(f"Output file: {output_excel}")
        print(f"Portal URL: {uhc_login_url}")
        print("=" * 50)
        
        # Run automation
        bot.run_automation(input_excel, output_excel, uhc_login_url)
        
    except KeyboardInterrupt:
        print("\n⚠️  Automation interrupted by user")
    except Exception as e:
        print(f"\n❌ An error occurred: {e}")
    finally:
        input("\nPress Enter to close the browser...")
        bot.close()