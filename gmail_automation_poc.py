from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
import traceback
import random
import logging
import os
from pathlib import Path

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('gmail_automation.log'),
        logging.StreamHandler()
    ]
)

class GmailAutomation:
    def __init__(self):
        self.driver = None
        self.logger = logging.getLogger(__name__)
        self.excel_path = "test_data.xlsx"
    
    def random_delay(self, min_seconds=1, max_seconds=3):
        delay = random.uniform(min_seconds, max_seconds)
        time.sleep(delay)
        return delay
    
    def setup_driver(self):
        try:
            self.logger.info("Setting up Chrome driver...")
            options = webdriver.ChromeOptions()
            
            # Updated Chrome options
            options.add_argument('--disable-gpu')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--start-maximized')
            options.add_argument('--disable-blink-features=AutomationControlled')
            options.add_experimental_option('excludeSwitches', ['enable-automation'])
            options.add_experimental_option('useAutomationExtension', False)
            
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
            self.driver.set_page_load_timeout(30)
            
            self.logger.info("Chrome driver setup completed successfully")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to setup Chrome driver: {str(e)}")
            return False

    def wait_for_element(self, by, value, timeout=20, condition="present"):
        try:
            self.logger.debug(f"Waiting for element: {value}")
            conditions = {
                "present": EC.presence_of_element_located,
                "clickable": EC.element_to_be_clickable,
                "visible": EC.visibility_of_element_located
            }
            
            element = WebDriverWait(self.driver, timeout).until(
                conditions[condition]((by, value))
            )
            return element
            
        except TimeoutException:
            self.logger.warning(f"Timeout waiting for element: {value}")
            self.take_screenshot(f"timeout_{value.replace('/', '_')}")
            return None

    def take_screenshot(self, name):
        try:
            timestamp = time.strftime("%Y%m%d-%H%M%S")
            filename = f"screenshots/{name}_{timestamp}.png"
            os.makedirs("screenshots", exist_ok=True)
            self.driver.save_screenshot(filename)
            self.logger.info(f"Screenshot saved: {filename}")
        except Exception as e:
            self.logger.error(f"Failed to take screenshot: {str(e)}")

    def login(self, email, password):
        """Enhanced login with better selectors"""
        try:
            if not self.setup_driver():
                return False, "Failed to setup WebDriver"

            self.logger.info("Starting login process...")
            self.driver.get('https://gmail.com')
            self.random_delay(3, 5)
            
            # Enter email
            email_field = self.wait_for_element(
                By.CSS_SELECTOR, 
                "input[type='email']",
                condition="visible"
            )
            if not email_field:
                return False, "Email field not found"
            
            email_field.clear()
            email_field.send_keys(email)
            self.random_delay(1, 2)
            email_field.send_keys(Keys.RETURN)
            self.random_delay(3, 4)
            
            # Enter password
            password_field = self.wait_for_element(
                By.CSS_SELECTOR, 
                "input[type='password']",
                condition="visible"
            )
            if not password_field:
                return False, "Password field not found"
            
            password_field.clear()
            password_field.send_keys(password)
            self.random_delay(1, 2)
            password_field.send_keys(Keys.RETURN)
            self.random_delay(5, 7)  # Increased delay for Gmail to load
            
            # Multiple checks for successful login
            compose_selectors = [
                "div[role='button'][gh='cm']",
                "div.T-I.T-I-KE.L3",
                "div.compose"
            ]
            
            for selector in compose_selectors:
                compose_button = self.wait_for_element(
                    By.CSS_SELECTOR, 
                    selector,
                    timeout=5,
                    condition="clickable"
                )
                if compose_button:
                    return True, "Success"
            
            return False, "Failed to find compose button"

        except Exception as e:
            self.logger.error(f"Login error: {str(e)}")
            return False, f"Error: {str(e)}"

    def compose_and_send_email(self, to_email, subject, body, attachment_path):
        """Enhanced email composition with better selectors and error handling"""
        try:
            # Find and click compose button with multiple selectors
            compose_selectors = [
                "div[role='button'][gh='cm']",
                "div.T-I.T-I-KE.L3",
                "div.compose"
            ]
            
            compose_clicked = False
            for selector in compose_selectors:
                compose_btn = self.wait_for_element(
                    By.CSS_SELECTOR, 
                    selector,
                    timeout=5,
                    condition="clickable"
                )
                if compose_btn:
                    compose_btn.click()
                    compose_clicked = True
                    break
            
            if not compose_clicked:
                return False, "Could not find compose button"
            
            self.random_delay(2, 3)
            
            # Enter recipient with multiple selectors
            to_selectors = [
                "textarea[name='to']",
                "input[role='combobox']",
                "input.agP.aFw"
            ]
            
            to_field = None
            for selector in to_selectors:
                to_field = self.wait_for_element(
                    By.CSS_SELECTOR, 
                    selector,
                    timeout=5,
                    condition="visible"
                )
                if to_field:
                    break
                    
            if not to_field:
                return False, "Could not find recipient field"
                
            to_field.send_keys(to_email)
            self.random_delay(1, 2)
            
            # Enter subject
            subject_selectors = [
                "input[name='subjectbox']",
                "input[aria-label='Subject']",
                "input.aoT"
            ]
            
            subject_field = None
            for selector in subject_selectors:
                subject_field = self.wait_for_element(
                    By.CSS_SELECTOR, 
                    selector,
                    timeout=5,
                    condition="visible"
                )
                if subject_field:
                    break
                    
            if not subject_field:
                return False, "Could not find subject field"
                
            subject_field.send_keys(subject)
            self.random_delay(1, 2)
            
            # Enter body with multiple selectors
            body_selectors = [
                "div[role='textbox']",
                "div.Am.Al.editable",
                "div.Ar.Au"
            ]
            
            body_field = None
            for selector in body_selectors:
                body_field = self.wait_for_element(
                    By.CSS_SELECTOR, 
                    selector,
                    timeout=5,
                    condition="visible"
                )
                if body_field:
                    break
                    
            if not body_field:
                return False, "Could not find email body field"
                
            body_field.send_keys(body)
            self.random_delay(1, 2)
            
            # Click attachment button and handle file upload
            attachment_selectors = [
                "div[command='Files']",
                "div[aria-label='Attach files']",
                "div.a1.aaA.aMZ"
            ]
            
            attach_button = None
            for selector in attachment_selectors:
                attach_button = self.wait_for_element(
                    By.CSS_SELECTOR, 
                    selector,
                    timeout=5,
                    condition="clickable"
                )
                if attach_button:
                    break
                    
            if not attach_button:
                return False, "Could not find attachment button"
                
            # Use hidden file input for upload
            file_input = self.driver.find_element(By.CSS_SELECTOR, "input[type='file']")
            file_input.send_keys(attachment_path)
            self.random_delay(5, 7)  # Wait for upload
            
            # Click send with multiple selectors
            send_selectors = [
                "div[role='button'][aria-label*='Send']",
                "div[data-tooltip='Send']",
                "div.T-I.J-J5-Ji.aoO.v7.T-I-atl.L3"
            ]
            
            send_clicked = False
            for selector in send_selectors:
                send_btn = self.wait_for_element(
                    By.CSS_SELECTOR, 
                    selector,
                    timeout=5,
                    condition="clickable"
                )
                if send_btn:
                    send_btn.click()
                    send_clicked = True
                    break
            
            if not send_clicked:
                return False, "Could not find send button"
                
            self.random_delay(3, 4)
            return True, "Email sent successfully"
            
        except Exception as e:
            self.logger.error(f"Error sending email: {str(e)}")
            return False, f"Failed to send email: {str(e)}"

    def logout(self):
        """Enhanced logout with better selectors"""
        try:
            # Multiple selectors for account button
            account_selectors = [
                "img[alt='Google Account']",
                "a[role='button'][aria-label*='Google Account']",
                "a.gb_d.gb_La.gb_f"
            ]
            
            account_clicked = False
            for selector in account_selectors:
                account_btn = self.wait_for_element(
                    By.CSS_SELECTOR, 
                    selector,
                    timeout=5,
                    condition="clickable"
                )
                if account_btn:
                    account_btn.click()
                    account_clicked = True
                    break
            
            if not account_clicked:
                return False
            
            self.random_delay(2, 3)
            
            # Click sign out with multiple approaches
            signout_selectors = [
                "div[data-action='sign out']",
                "a[text='Sign out']",
                "a.gb_Cb.gb_Vf.gb_4f.gb_Re.gb_4c"
            ]
            
            for selector in signout_selectors:
                try:
                    signout_btn = self.wait_for_element(
                        By.CSS_SELECTOR, 
                        selector,
                        timeout=5,
                        condition="clickable"
                    )
                    if signout_btn:
                        signout_btn.click()
                        return True
                except:
                    continue
            
            # Try finding element by text if CSS selectors fail
            try:
                signout_btn = self.driver.find_element(By.XPATH, "//*[contains(text(), 'Sign out')]")
                signout_btn.click()
                return True
            except:
                pass
            
            return False
            
        except Exception as e:
            self.logger.error(f"Logout error: {str(e)}")
            return False

def main():
    try:
        # Read test data from Excel
        df = pd.read_excel("test_data.xlsx")
        
        automation = GmailAutomation()
        
        for index, row in df.iterrows():
            # Perform login
            success, message = automation.login(
                email=row['Email'],
                password=row['Password']
            )
            
            # Update Excel with login results
            df.at[index, 'Login_Status'] = message
            
            if success:
                # Send email if login successful
                email_success, email_message = automation.compose_and_send_email(
                    to_email=row['To_Email'],
                    subject=row['Subject'],
                    body=row['Body'],
                    attachment_path=row['Attachment_Path']
                )
                df.at[index, 'Email_Status'] = email_message
                
                # Logout
                if automation.logout():
                    df.at[index, 'Logout_Status'] = "Success"
                else:
                    df.at[index, 'Logout_Status'] = "Failed"
            
            # Save results back to Excel
            df.to_excel("test_data.xlsx", index=False)
            
            if automation.driver:
                automation.driver.quit()
                
    except Exception as e:
        logging.error(f"Error in main: {str(e)}")
        traceback.print_exc()

if __name__ == "__main__":
    
    main()