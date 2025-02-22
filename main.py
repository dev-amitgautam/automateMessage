from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from concurrent.futures import ThreadPoolExecutor
import threading
import queue
import logging
import time
import datetime
from dotenv import load_dotenv
import os

load_dotenv()

def create_driver():
    """Create and return a new Firefox driver instance"""
    gecko_driver_path = r'C:\Users\amitg\OneDrive\Desktop\Selenium\automateMessage\geckodriver.exe'
    service = Service(gecko_driver_path)
    driver = webdriver.Firefox(service=service)
    return driver

# Open the login page
# driver.get('https://kudos.kotak.com/webint/app/login/login')  # Replace with your login page URL

# Wait for the page to load
# time.sleep(2)

# crn_numbers = ["590445807", "482245420"]  # Replace with your actual CRN numbers
userName = os.getenv("UNAME")
password = os.getenv("PASS")

print(userName)
print(password)

def process_single_crn(crn_number, username, password, result_queue):

    driver = create_driver()

    def click_lead_name_link():
        """Find and click the lead name link"""
        try:
            # Wait for page to be fully loaded
            WebDriverWait(driver, 20).until(
                lambda driver: driver.execute_script('return document.readyState') == 'complete'
            )
            time.sleep(1)  # Additional wait to ensure dynamic content is loaded
            
            # Try to find the lead name link
            lead_link = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'a.tb.button-icon-text.f12[data-autoid^="LeadName_"]'))
            )
            print(lead_link)
            if not lead_link:
                print("Message already sent - Lead name link not found")
                return True

            # Scroll the link into view
            driver.execute_script("arguments[0].scrollIntoView(true);", lead_link)
            time.sleep(1)
            
            # Try to click using JavaScript for better reliability
            try:
                lead_link.click()
            except:
                driver.execute_script("arguments[0].click();", lead_link)
                
            print("Successfully clicked lead name link")
            time.sleep(2)  # Wait for page load

            # Look for and click the LOP Link for CCE
            try:
                lop_link = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'a.tb.button-icon-text.f12.button-icon-text--icon[title="LOP Link for CCE"]'))
                )
                
                # Scroll the link into view
                driver.execute_script("arguments[0].scrollIntoView(true);", lop_link)
                time.sleep(1)
                
                # Try to click using JavaScript for better reliability
                try:
                    lop_link.click()
                except:
                    driver.execute_script("arguments[0].click();", lop_link)
                    
                print("Successfully clicked LOP Link for CCE")
                time.sleep(2)  # Wait for popup to load
            except Exception as e:
                print(f"Error clicking LOP Link for CCE: {str(e)}")

            # need to click a tag to open new pop -up

            windows = driver.window_handles
            if len(windows) > 1:
                driver.switch_to.window(windows[-1])
                print("Switched to eligibility check window")

            # Wait for and click the Next button
            try:
                WebDriverWait(driver, 20).until(
                    lambda driver: driver.execute_script('return document.readyState') == 'complete'
                )
                # Additional wait for any dynamic content
                time.sleep(2)

                # Now wait for the next button
                next_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-testid="button1"]'))
                )
                
                # Scroll the Next button into view and click
                driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
                time.sleep(1)
                
                driver.execute_script("arguments[0].click();", next_button)
                time.sleep(2)  # Wait for the new page to load
                
                # Wait for and fill in the CRMNext User ID field
                try:
                    time.sleep(3)  # Add more wait time for the field to be fully loaded
                    
                    # Wait for page to stabilize and try multiple times if needed
                    max_attempts = 3
                    attempt = 0
                    
                    while attempt < max_attempts:
                        try:
                            # Try to find the input using multiple selectors with a fresh wait
                            user_id_input = WebDriverWait(driver, 15).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, 'input[id="txtmissing_563"]'))
                            )
                            
                            # Ensure element is interactable
                            WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[id="txtmissing_563"]'))
                            )
                            
                            # Scroll the input into view
                            driver.execute_script("arguments[0].scrollIntoView(true);", user_id_input)
                            time.sleep(1)
                            
                            # Try multiple methods to enter the text
                            try:
                                # user_id_input.clear()
                                user_id_input.send_keys(userName)
                                # Verify the value was entered
                                print("user_id_input.get_attribute('value')", user_id_input.get_attribute('value'))
                                if user_id_input.get_attribute('value') == userName:
                                    print("Successfully entered CRMNext User ID", crn_number)


                                    submit_button = WebDriverWait(driver, 10).until(
                                        EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[id="missing_632"]'))
                                    )

                                    # Scroll the button into view
                                    driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)
                                    time.sleep(1)  # Wait for the scroll to complete
                                    
                                    # Try to click using JavaScript for better reliability
                                    try:
                                        submit_button.click()
                                    except:
                                        driver.execute_script("arguments[0].click();", submit_button)
                                    
                                    print("Successfully clicked Submit button")
                                    time.sleep(10)
                                    
                                    # Close the current window (popup)
                                    driver.close()
                                    
                                    # Switch back to the main window
                                    driver.switch_to.window(driver.window_handles[0])
                                    
                                    # Refresh the main page
                                    driver.refresh()

                                    time.sleep(3)
                                    
                                    # Close the main window and quit the driver
                                    try:
                                        driver.close()  # Close the current window
                                        driver.quit()   # Quit the driver and close all associated windows
                                        print("Successfully closed browser and ended test")
                                        return True
                                    except Exception as e:
                                        print(f"Error while closing browser: {str(e)}")
                                        exit(1)  


                            except:
                                print("Failed to enter CRMNext User ID")
                            
                            attempt += 1
                            if attempt == max_attempts:
                                print("Failed to enter CRMNext User ID after multiple attempts")
                            
                        except Exception as inner_e:
                            print(f"Attempt {attempt + 1} failed: {str(inner_e)}")
                            attempt += 1
                            time.sleep(2)
                except Exception as e:
                    print(f"Error entering CRMNext User ID: {str(e)}")
            except Exception as e:
                print(f"Error clicking Next button: {str(e)}")


        except Exception as e:
            print(f"Error clicking lead name link: {str(e)}")
            return False


    try:
        # Wait for the page to load
        time.sleep(2)

        driver.get('https://kudos.kotak.com/webint/app/login/login')

        username_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'TxtName'))
        )

        # Wait for the password field to be present
        password_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'TxtPassword'))
        )

        # Wait for the login button to be present
        login_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input.login-button[type="submit"]'))
        )

        # Enter your credentials
        username_field.send_keys(userName)  # Replace with your actual username
        password_field.send_keys(password)  # Replace with your actual password

        # Click the login button
        login_button.click()

        # Wait for the login process to complete
        time.sleep(3)

        # Check if login was successful (e.g., by checking for a specific element on the logged-in page)
        try:
            logged_in_element = driver.find_element(By.ID, 'logged-in-element-id')  # Replace with an element that appears after login
            print("Login successful!")
        except:
            print("Login failed.")

        time.sleep(2)  # Adjust the sleep time as needed

        # Locate the Offers link in the side navigation menu
        offers_link = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//a[@href="/webint/app/CRMNextObject/Home/Offer"]'))
        )

        # Click the Offers link
        offers_link.click()

        # Wait for the Offers page to load
        time.sleep(3)

        # Check if the Offers page was loaded successfully
        try:
            offers_page_element = driver.find_element(By.ID, 'offers-page-element-id')  # Replace with an element that appears on the Offers page
            print("Offers page loaded successfully!")
        except:
            print("Failed to load the Offers page.")

        # Wait for the search icon to be present
        search_icon = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'a.field__item.input-icon.pointer'))
        )

        # Click the search icon
        search_icon.click()

        # Wait for the search functionality to load
        time.sleep(2)

        crn_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, 'cust_418'))
            )

        # Clear the input field (if needed) and enter the CRN number
        crn_input.clear()
        crn_input.send_keys(crn_number)

        # Wait for the Search button to be present
        search_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'a[data-autoid="Search"]'))
        )

        # Click the Search button
        search_button.click()

        # Wait for the search results to load
        time.sleep(5)

        # Check if the search was successful (e.g., by checking for a specific element in the search results)
        try:
            search_results = driver.find_element(By.ID, 'search-results-container')  # Replace with an element that appears in the search results
            print(f"Search for CRN {crn_number} successful!")
        except:
            print(f"Search for CRN {crn_number} failed.")

        def get_date_from_div(div):
            """Extract date and offer link div from a div element"""
            try:
                # Find the date element within the div
                date_element = div.find_element(By.CSS_SELECTOR, "div[data-autoid$='_CreatedOn_val']")
                date_text = date_element.get_attribute('title')
                
                # Parse the date (format: "DD/MM/YYYY HH:MM AM/PM")
                date_time = datetime.datetime.strptime(date_text, "%d/%m/%Y %I:%M %p")
                return date_time, div
            except Exception as e:
                print(f"Error extracting date: {str(e)}")
                return None, None

        def find_latest_date_and_click():
            """Find the latest date from February 2025 and click its offer"""
            latest_date = None
            latest_div = None
            page_count = 1
            
            while True:
                try:
                    # Wait for offer divs to be present
                    offer_divs = WebDriverWait(driver, 10).until(
                        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.flex-1.pt-6.ph-8"))
                    )
                    
                    print(f"Checking page {page_count}")
                    
                    # First define the reference date at the top of the script, after the imports
                    REFERENCE_DATE = datetime.datetime(2025, 2, 10, 00, 00)  # 5:55 PM on Jan 20, 2025

                    # Then modify the date checking logic in the find_latest_date_and_click function
                    for div in offer_divs:
                        date, current_div = get_date_from_div(div)
                        if date and date >= REFERENCE_DATE:
                            if latest_date is None or date > latest_date:
                                latest_date = date
                                latest_div = current_div
                    
                    # If we found a date in January 2025, click the offer
                    if latest_div:
                        print(f"Found latest date: {latest_date}")
                        offer_link = latest_div.find_element(By.CSS_SELECTOR, "a[data-autoid$='_OfferNo_val']")
                        # Scroll the offer link into view before clicking
                        driver.execute_script("arguments[0].scrollIntoView(true);", offer_link)
                        time.sleep(1)
                        offer_link.click()

                        try:
                            check_eligibility = WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, 'div[data-autoid="webinturl_25_ctrl"] a'))
                            )
                            
                            # Scroll the Check Eligibility button into view and click
                            driver.execute_script("arguments[0].scrollIntoView(true);", check_eligibility)
                            time.sleep(1)  # Wait for scroll to complete
                            
                            # Try to click using JavaScript if regular click fails
                            try:
                                check_eligibility.click()
                            except:
                                driver.execute_script("arguments[0].click();", check_eligibility)
                                
                            # Continue with existing eligibility check flow...
                            
                        except:
                            print("Check Eligibility button not found, looking for lead name link...")
                            if click_lead_name_link():
                                return True
                        
                        # Wait for the popup window to appear
                        time.sleep(2)
                        
                        # Switch to the new window if one opens
                        windows = driver.window_handles
                        if len(windows) > 1:
                            driver.switch_to.window(windows[-1])
                            print("Switched to eligibility check window")
                        
                        # Wait for and handle the "Proceed without Cover" checkbox if it exists
                        try:
                            # Wait for all radio button containers
                            radio_containers = WebDriverWait(driver, 10).until(
                                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.viv_radioInlineElementKotak'))
                            )
                            
                            # Find the container with "Proceed without Cover"
                            for container in radio_containers:
                                try:
                                    label = container.find_element(By.CSS_SELECTOR, 'label.viv_radioLabelKotak')
                                    if "Proceed without Cover" in label.text:
                                        # Click the label using JavaScript for better reliability
                                        driver.execute_script("arguments[0].click();", label)
                                        print("Successfully clicked 'Proceed without Cover' radio button")
                                        time.sleep(2)  # Wait for any UI updates
                                        break
                                except Exception as inner_e:
                                    print(f"Error processing radio container: {str(inner_e)}")
                                    continue

                        except Exception as e:
                            print(f"Error with radio button selection: {str(e)}")
                            print("Proceed without Cover option not found or not clickable, continuing...")

                        # Wait for and handle the loan purpose dropdown
                        try:
                            time.sleep(2)  # Add extra wait time for dropdown to be fully loaded
                            
                            # Wait for the dropdown container and dropdown with the correct ID
                            loan_purpose_dropdown = WebDriverWait(driver, 15).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, '#missing_2409 select#ddl_undefined'))
                            )
                            
                            # Ensure element is interactable
                            WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, '#missing_2409 select#ddl_undefined'))
                            )
                            
                            # Scroll the dropdown into view
                            driver.execute_script("arguments[0].scrollIntoView(true);", loan_purpose_dropdown)
                            time.sleep(1)
                            
                            # Select "Personal Expense" from dropdown
                            select = Select(loan_purpose_dropdown)
                            select.select_by_value("PE")
                            time.sleep(1)  # Wait for selection to take effect
                        except Exception as e:
                            print(f"Error with dropdown: {str(e)}")
                        
                        # Wait for and click the Next button
                        try:
                            next_button = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-testid="button1"]'))
                            )
                            
                            # Scroll the Next button into view and click
                            driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
                            time.sleep(1)
                            
                            driver.execute_script("arguments[0].click();", next_button)
                            time.sleep(2)  # Wait for the new page to load
                            
                            # Wait for and fill in the CRMNext User ID field
                            try:
                                time.sleep(3)  # Add more wait time for the field to be fully loaded
                                
                                # Wait for page to stabilize and try multiple times if needed
                                max_attempts = 3
                                attempt = 0
                                
                                while attempt < max_attempts:
                                    try:
                                        # Try to find the input using multiple selectors with a fresh wait
                                        user_id_input = WebDriverWait(driver, 15).until(
                                            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[data-testid="textbox"]'))
                                        )
                                        
                                        # Ensure element is interactable
                                        WebDriverWait(driver, 10).until(
                                            EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[data-testid="textbox"]'))
                                        )
                                        
                                        # Scroll the input into view
                                        driver.execute_script("arguments[0].scrollIntoView(true);", user_id_input)
                                        time.sleep(1)
                                        
                                        # Try multiple methods to enter the text
                                        try:
                                            user_id_input.clear()
                                            user_id_input.send_keys(userName)
                                            # Verify the value was entered
                                            print("user_id_input.get_attribute('value')", user_id_input.get_attribute('value'))
                                            if user_id_input.get_attribute('value') == userName:
                                                print("Successfully entered CRMNext User ID")


                                                submit_button = WebDriverWait(driver, 10).until(
                                                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[id="missing_26865"][data-testid="button1"]'))
                                                )

                                                # Scroll the button into view
                                                driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)
                                                time.sleep(1)  # Wait for the scroll to complete

                                                # Try to click using JavaScript for better reliability
                                                try:
                                                    submit_button.click()
                                                except:
                                                    driver.execute_script("arguments[0].click();", submit_button)

                                                print("Successfully clicked Submit button")
                                                time.sleep(10)
                                                
                                                # Close the current window (popup)
                                                driver.close()
                                                
                                                # Switch back to the main window
                                                driver.switch_to.window(driver.window_handles[0])
                                                
                                                # Refresh the main page
                                                driver.refresh()
                                                
                                                # Wait for the page to reload
                                                time.sleep(3)

                                                # Wait for and click the lead name link
                                                if click_lead_name_link():
                                                    return True

                                        except:

                                            print("Failed to enter CRMNext User ID")
                                        
                                        attempt += 1
                                        if attempt == max_attempts:
                                            print("Failed to enter CRMNext User ID after multiple attempts")
                                        
                                    except Exception as inner_e:
                                        print(f"Attempt {attempt + 1} failed: {str(inner_e)}")
                                        attempt += 1
                                        time.sleep(2)
                                    
                            except Exception as e:
                                print(f"Error entering CRMNext User ID: {str(e)}")
                        except Exception as e:
                            print(f"Error clicking next button: {str(e)}")
                        
                        return True
                    
                    # If no date found, try next page
                    next_button = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'a[data-autoid="nextButton_CrmGrid"]'))
                    )
                    
                    # Check if next button is disabled
                    if 'disabled' in next_button.get_attribute('class'):
                        print("Reached last page without finding January 2025 date")
                        break
                    
                    # Scroll the next button into view and click
                    driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
                    time.sleep(1)  # Wait for scroll to complete
                    
                    # Try to click using JavaScript if regular click fails
                    try:
                        next_button.click()
                    except:
                        driver.execute_script("arguments[0].click();", next_button)
                    
                    page_count += 1
                    time.sleep(2)  # Wait for page to load
                    
                except Exception as e:
                    print(f"Error on page {page_count}: {str(e)}")
                    break
            
            return False

        # Replace the existing date finding code with this call
        try:
            if find_latest_date_and_click():
                print("Successfully found and clicked January 2025 offer")
                result_queue.put({"crn": crn_number, "status": "success"})
            else:
                print("No February 2025 offers found")
        except Exception as e:
            print(f"Error during search: {str(e)}")

    except Exception as e:
        logging.error(f"Error processing CRN {crn_number}: {str(e)}")
        result_queue.put({"crn": crn_number, "status": "failed", "error": str(e)})
    finally:
        if driver:
            try:
                driver.quit()
            except:
                pass


# finally:
#     # Close the browser
#     driver.quit()

def run_parallel_tests():
    crn_numbers = ["707130140", "720694715", "683065677", "683142443"]  # Add more CRN numbers as needed
    max_workers = min(len(crn_numbers), 4)  # Limit concurrent threads
    result_queue = queue.Queue()
    
    logging.basicConfig(level=logging.INFO)
    
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [
            executor.submit(
                process_single_crn, 
                crn, 
                os.getenv("UNAME"), 
                os.getenv("PASS"),
                result_queue
            ) for crn in crn_numbers
        ]
        
        # Wait for all tasks to complete
        for future in futures:
            try:
                future.result(timeout=600)  # 10-minute timeout per thread
            except Exception as e:
                logging.error(f"Thread execution failed: {str(e)}")
    
    # Process results
    results = []
    while not result_queue.empty():
        results.append(result_queue.get())
    
    # Print summary
    success_count = sum(1 for r in results if r["status"] == "success")
    logging.info(f"Processed {len(results)} CRNs. Success: {success_count}, Failed: {len(results) - success_count}")

if __name__ == "__main__":
    run_parallel_tests()