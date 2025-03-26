import os
import re
import time
import logging
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
class SharePoint:
    def __init__(
        self,
        username: str,
        password: str,
        timeout: int = 10,
        headless:bool=False,
        download_directory:str = os.path.dirname(os.path.abspath(__file__)),
        logger_name: str = __name__,
    ):
        os.makedirs(download_directory,exist_ok=True)
        options = webdriver.ChromeOptions()
        options.add_argument('--disable-notifications')
        options.add_experimental_option("prefs", {
            "download.default_directory": download_directory,  
            "download.prompt_for_download": False, 
            "safebrowsing.enabled": True 
        })
        # Disable log
        options.add_argument('--disable-logging')
        options.add_argument('--log-level=3')  
        options.add_argument('--silent')
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        if headless: 
            options.add_argument('--headless=new')
        # Attribute
        self.logger = logging.getLogger(logger_name)
        self.browser = webdriver.Chrome(options=options)
        self.browser.maximize_window()
        self.timeout = timeout
        self.wait = WebDriverWait(self.browser, timeout)
        self.username = username
        self.password = password
        # Trạng thái đăng nhập
        self.authenticated = self.__authentication(username, password)

    def __authentication(self, username: str, password: str) -> bool:
        time.sleep(0.5)
        self.browser.get("https://login.microsoftonline.com/")
        # -- Wait load page
        self.wait.until(
            lambda driver: driver.execute_script("return document.readyState") == "complete"
        )
        if self.browser.current_url == "https://m365.cloud.microsoft/?auth=2":
            return True
        try:
            # -- Email, phone or Skype
            self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR,"input[id=\"i0116\"]")
                )
            ).send_keys(username)
            time.sleep(0.5)
            # -- Next
            self.wait.until(
                EC.presence_of_element_located(
                    (By.ID,"idSIButton9")
                )
            ).click()
            time.sleep(1)
            # -- Check usernameError
            try:
                usernameError = self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR,"div[id=\"usernameError\"]")
                    )
                )
                self.logger.error(usernameError.text)
                return False
            except TimeoutException:
                pass
            # -- Password
            self.wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR,"input[id=\"i0118\"]")
                )
            ).send_keys(password)
            time.sleep(0.5)
            # -- Sign in
            self.wait.until(
                EC.presence_of_element_located(
                    (By.ID,"idSIButton9")
                )
            ).click()
            time.sleep(0.5)
            # -- Check stay signed in
            try:
                passwordError = self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR,"div[id=\"passwordError\"]")
                    )
                )
                self.logger.error(passwordError.text)
                return False
            except TimeoutException:
                pass
            # -- Stay signed in?
            try:
                self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR,"div[class='row text-title']")
                    )
                )
                self.wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR,"input[id='idSIButton9']")
                    )
                ).click()
                time.sleep(0.5)
                self.wait.until(
                    lambda driver: driver.execute_script("return document.readyState") == "complete"
                )
                if self.browser.current_url.startswith("https://m365.cloud.microsoft/?auth="):
                    self.logger.info("Xác thực thành công")
                    return True
                else:
                    return False
            except TimeoutException as e:
                self.logger.error(e)
                return False
        except TimeoutException as e:
            self.logger.error(e)
            return False
        except Exception as e:
            self.logger.error(e)
            return False
        
    def download_file(self,site_url:list[str],file_pattern:str) -> bool: 
        result = []
        for url in site_url:
            try:
                time.sleep(0.5)
                self.browser.get(url)
                time.sleep(2.5)
                # -- Folder --
                folders = file_pattern.split("/")[:-1]
                for folder in folders:
                    # Found folder
                    ms_DetailsList_contentWrapper = self.wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR,"div[class='ms-DetailsList-contentWrapper']"))
                    )
                    rows = ms_DetailsList_contentWrapper.find_elements(
                        by = By.CSS_SELECTOR,
                        value = "div[class^='ms-DetailsRow-fields fields-']"
                    ) # Lấy tất cả các dòng 
                    for row in rows:
                        icon_gridcell = row.find_element(
                            By.CSS_SELECTOR, 
                            "div[role='gridcell'][data-automationid='DetailsRowCell']"
                        )
                        if icon_gridcell.find_elements(By.TAG_NAME, "svg"): # Folder
                            name_gridcell = row.find_element(
                            By.CSS_SELECTOR, 
                            "div[role='gridcell'][data-automation-key^='displayNameColumn_']"
                            )
                            button = name_gridcell.find_element(By.TAG_NAME,'button')
                            if button.text == folder:
                                button.click()
                                time.sleep(5)
                                break
                            
                # -- File --
                pattern = file_pattern.split("/")[-1]
                pattern = re.compile(pattern)
                ms_DetailsList_contentWrapper = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR,"div[class='ms-DetailsList-contentWrapper']"))
                )       
                rows = ms_DetailsList_contentWrapper.find_elements(
                    by = By.CSS_SELECTOR,
                    value = "div[class^='ms-DetailsRow-fields fields-']"
                ) # Lấy tất cả các dòng  
                for row in rows:
                    icon_gridcell = row.find_element(
                        By.CSS_SELECTOR, 
                        "div[role='gridcell'][data-automationid='DetailsRowCell']"
                    )
                    if icon_gridcell.find_elements(By.TAG_NAME, "svg"): # Folder
                        continue
                    else:
                        name_gridcell = row.find_element(
                            By.CSS_SELECTOR, 
                            "div[role='gridcell'][data-automation-key^='displayNameColumn_']"
                        )
                        button = name_gridcell.find_element(By.TAG_NAME,'button')
                        display_name = button.text
                        # Nếu display_name match với file là được
                        if pattern.match(display_name):
                            gridcell_div = row.find_element(By.XPATH, "./preceding-sibling::div[@role='gridcell']")
                            checkbox = gridcell_div.find_element(By.CSS_SELECTOR, "div[role='checkbox']")
                            if checkbox.get_attribute('aria-checked') == "false": 
                                checkbox.click()
                # Download
                try:
                    self.wait.until(
                        EC.presence_of_element_located((By.XPATH, "//span[text()='Download']"))
                    )
                    self.wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//span[text()='Download']"))
                    ).click()
                except TimeoutException:
                    self.wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR,"i[data-icon-name='More']"))
                    )
                    self.wait.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR,"i[data-icon-name='More']"))
                    ).click()
                    self.wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "button[name='Download']"))
                    )
                    self.wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "button[name='Download']"))
                    ).click()
                time.sleep(0.5)
                result.append(True)
            except Exception as e:
                self.logger.error(e)
                result.append(False)
        time.sleep(10)
        return result
            

            
                
                
                
            