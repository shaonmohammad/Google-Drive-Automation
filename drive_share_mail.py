import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import threading
import time
import random
import os

class AutomationGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Google Drive Automation")
        self.root.geometry("800x600")
        
        frame = tk.Frame(self.root, padx=10, pady=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(frame, text="Emails per batch:").pack()
        self.emails_per_batch = tk.Entry(frame)
        self.emails_per_batch.pack()
        
        tk.Label(frame, text="Runs per email:").pack()
        self.runs_per_email = tk.Entry(frame)
        self.runs_per_email.pack()
        
        self.select_button = tk.Button(frame, text="Select File", command=self.select_file, width=20)
        self.select_button.pack(pady=5)
        
        self.admin_button = tk.Button(frame, text="Select Admin Emails", command=self.select_admin_file, width=20)
        self.admin_button.pack(pady=5)
        
        self.client_button = tk.Button(frame, text="Select Client Emails", command=self.select_client_file, width=20)
        self.client_button.pack(pady=5)
        
        self.log_area = scrolledtext.ScrolledText(frame, height=15)
        self.log_area.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.start_btn = tk.Button(frame, text="Start", command=self.start_automation)
        self.start_btn.pack(pady=5)

    def select_file(self):
        self.file_name = os.path.basename(filedialog.askopenfilename(
            title="Select Input File",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
        ))
        return self.file_name

    def select_admin_file(self):
        self.admin_file = filedialog.askopenfilename(
            title="Select Admin Emails File",
            filetypes=[("Text Files", "*.txt")]
        )
        if self.admin_file:
            self.log_message(f"Selected admin file: {os.path.basename(self.admin_file)}")

    def select_client_file(self):
        self.client_file = filedialog.askopenfilename(
            title="Select Client Emails File",
            filetypes=[("Text Files", "*.txt")]
        )
        if self.client_file:
            self.log_message(f"Selected client file: {os.path.basename(self.client_file)}")

    def log_message(self, msg):
        self.log_area.insert(tk.END, f"{msg}\n")
        self.log_area.see(tk.END)

    def start_automation(self):
        try:
            if not hasattr(self, 'admin_file') or not hasattr(self, 'client_file') or not hasattr(self, 'file_name'):
                self.log_message("Please select all required files")
                return
                
            emails = int(self.emails_per_batch.get())
            runs = int(self.runs_per_email.get())

            self.start_btn.config(state=tk.DISABLED)
            thread = threading.Thread(target=self.run_automation, args=(emails, runs))
            thread.daemon = True
            thread.start()
            
        except ValueError:
            self.log_message("Please enter valid numbers")
            
    def run_automation(self, emails_per_batch, runs_per_email):
        try:
            automation = GoogleDriveAutomation(
                file_path=self.file_name,
                admin_file=self.admin_file,
                client_file=self.client_file,
                emails_per_batch=emails_per_batch,
                runs_per_email=runs_per_email,
                logger=self
            )
            automation.run()
        except Exception as e:
            self.log_message(f"Error: {str(e)}")
        finally:
            self.start_btn.config(state=tk.NORMAL)

class GoogleDriveAutomation:
    def __init__(self, file_path, admin_file, client_file, emails_per_batch, runs_per_email, logger):
        self.file_path = file_path
        self.admin_file = admin_file
        self.client_file = client_file
        self.emails_per_batch = emails_per_batch
        self.runs_per_email = runs_per_email
        self.logger = logger
        self.failed_emails = set()
        # Add output directory path
        self.output_dir = os.path.dirname(self.client_file)
        self.setup_driver()
        
        # Register exit handler
        import atexit
        atexit.register(self.save_failed_emails)

    def setup_driver(self):
        options = webdriver.ChromeOptions()
        options.add_argument('--start-maximized')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        self.driver = webdriver.Chrome(options=options)
        self.wait = WebDriverWait(self.driver, 20)

    def save_failed_emails(self):
        if self.failed_emails:
            failed_file = os.path.join(self.output_dir, "failed_mail.txt")
            with open(failed_file, "w") as f:
                for email in self.failed_emails:
                    f.write(f"{email}\n")
            # self.logger.log_message(f"Failed emails saved to: {failed_file}")

    def load_credentials(self):
        with open(self.admin_file, "r") as f:
            return [line.strip().split(":") for line in f.readlines()]

    def load_client_emails(self):
        with open(self.client_file, "r") as f:
            return [line.strip() for line in f.readlines()]
    
    def login(self, email, password):
        try:
            self.driver.get("https://accounts.google.com/v3/signin/identifier?continue=https%3A%2F%2Fdrive.google.com%2Fdrive%2F%3Fdmr%3D1%26ec%3Dwgc-drive-hero-goto&followup=https%3A%2F%2Fdrive.google.com%2Fdrive%2F%3Fdmr%3D1%26ec%3Dwgc-drive-hero-goto&ifkv=AcMMx-fKwqE5UVY-hlsXuXKxoofz3reh-BGOSm7tg2YgKZfjbLQm8D_HXc8U4Cie50Vrn1XwKRaqMg&osid=1&passive=1209600&service=wise&flowName=GlifWebSignIn&flowEntry=ServiceLogin&dsh=S906190291%3A1733673434469058&ddm=1")
            
            email_input = self.wait.until(EC.presence_of_element_located((By.ID, "identifierId")))
            email_input.send_keys(email, Keys.ENTER)
            
            password_input = self.wait.until(EC.presence_of_element_located((By.NAME, "Passwd")))
            password_input.send_keys(password, Keys.ENTER)
            
            try:
                wrong_password = self.wait.until(
                    EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Wrong password')]"))
                )
                if wrong_password:
                    self.logger.log_message(f"Wrong password for account: {email}")
                    self.failed_emails.add(email)
                    return False
            except:
                pass
            
            time.sleep(2)
            
            wait_5_sec = WebDriverWait(self.driver, 5)
            try:
                wait_5_sec.until(EC.url_contains("drive.google.com/drive/home"))
                self.logger.log_message(f"Logged in: {email}")
                return True
            except:
                self.logger.log_message(f"This is new account. Try again later after setup this account manually")
                self.failed_emails.add(email)
                return False
                
        except Exception as e:
            self.logger.log_message(f"Login failed with this email: {email}")
            self.failed_emails.add(email)
            return False

    def upload_and_share(self, batch_emails):
        try:
            file_element = WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located(
                    (By.XPATH, f"//div[@jsname='LvFR7c' and .//div[text()='{self.file_path}']]")
                )
            )

            self.driver.execute_script("arguments[0].scrollIntoView(true);", file_element)
            
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(file_element))
            file_element.click()

            share_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "[aria-label='Share Ctrl+Alt+A']"))
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true);", share_button)
            share_button.click()

            iframe = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "ea-Rc-x-Vc"))
            )
            self.driver.switch_to.frame(iframe)
            time.sleep(2)
            
            for email in batch_emails:
                self.driver.switch_to.active_element.send_keys(email)
                time.sleep(2)
                self.driver.switch_to.active_element.send_keys(Keys.ENTER)
                time.sleep(0.5)
            self.logger.log_message(f"Entered email")
            
            send_button = self.driver.find_element(By.CLASS_NAME, "UywwFc-LgbsSe")
            self.driver.execute_script("arguments[0].click();", send_button)
            time.sleep(3)
            self.driver.switch_to.active_element.send_keys(Keys.ENTER)
            self.logger.log_message(f"Successfully Shared with {len(batch_emails)} emails:{batch_emails}")
            self.driver.refresh()
            return True
            
        except Exception as e:
            self.logger.log_message(f"Failed to share file with this email batch")
            return False

    def run(self):
        try:
            while True:
                client_emails = self.load_client_emails()
                if not client_emails:
                    self.logger.log_message("No more client emails to process. Exiting...")
                    break
                    
                credentials = self.load_credentials()
                for email, password in credentials:
                    self.logger.log_message(f"Processing {email}")
                    
                    if not self.login(email, password):
                        continue
                    
                    new_btn = self.wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//button[.//span[text()='New']]")))
                    new_btn.click()
                    
                    upload_btn = self.wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//div[text()='File upload']")))
                    upload_btn.click()
                    
                    time.sleep(2)
                    os.system(f'powershell -Command "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait(\'{self.file_path}\'); [System.Windows.Forms.SendKeys]::SendWait(\'~\')"')
                    
                    try:
                        duplicate_upload = self.wait.until(EC.element_to_be_clickable(
                            (By.CLASS_NAME, "nCP5yc")))     
                        duplicate_upload.click()
                    except Exception as e:
                        pass

                    self.wait.until(EC.presence_of_element_located(
                        (By.XPATH, "//div[@aria-label='1 upload complete']")))
                    self.logger.log_message("File uploaded")
                    
                    self.driver.refresh()
                    time.sleep(2)

                    for _ in range(self.runs_per_email):
                        client_emails = self.load_client_emails()
                        if not client_emails:
                            break
                            
                        batch = client_emails[:self.emails_per_batch]
                        if self.upload_and_share(batch):
                            with open(self.client_file, "w") as f:
                                f.writelines(f"{e}\n" for e in client_emails[self.emails_per_batch:])
                        
                        time.sleep(1)
                    
                    self.driver.delete_all_cookies()
                    self.driver.refresh()
                    
                self.logger.log_message("Completed all credentials. Running again for remaining client emails...")
                
        finally:
            self.save_failed_emails()
            self.driver.quit()

if __name__ == "__main__":
    app = AutomationGUI()
    app.root.mainloop()