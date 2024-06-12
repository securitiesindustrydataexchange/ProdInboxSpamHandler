import argparse
import logging
import os
from datetime import datetime

from selenium import webdriver
from selenium.common import TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

# Configure logging
project_root = os.path.dirname(os.path.abspath(__file__))
current_date = datetime.now().strftime('%Y%m%d')
log_file_name = f'prod_inbox_spam_handler_{current_date}.log'
log_file_path = os.path.join(project_root, log_file_name)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', handlers=[
        logging.FileHandler(log_file_path),
        logging.StreamHandler()
    ])
logger = logging.getLogger(__name__)

# Configure Chrome options
chrome_options = Options()
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--incognito")

# Automatically download the appropriate version of ChromeDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)


# Outlook login
def login_to_outlook(email, password):
    driver.get("https://outlook.office.com")
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.NAME, "loginfmt"))
    ).send_keys(email + Keys.RETURN)
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.NAME, "passwd"))
    ).send_keys(password + Keys.RETURN)
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "idBtn_Back"))
    ).click()


def is_spam(input_string):
    error_messages = [
        'host2132.hostmonster.com rejected your message',
        "Your message to datafeed.prod@sidexchangedemo.com couldn't be delivered",
        "Your message to datafeed2.prod@sidexchangedemo.com couldn't be delivered"
    ]

    for error_message in error_messages:
        if error_message in input_string:
            return True
    return False


# monitor
def monitor_spam():
    processed_email = []
    logger.info("Processing SPAM emails")
    while True:
        logging.info('Monitoring...')
        try:
            email_list = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, "//*[@id='MailList']/div/div/div/div/div/div/div/div/div/div"))
            )
            for an_email in email_list:
                email_id = an_email.get_attribute("data-convid")
                if email_id in processed_email:
                    continue
                email_body = an_email.get_attribute('aria-label')
                if is_spam(email_body):
                    WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable(
                            (By.CSS_SELECTOR, "[data-convid='{}']".format(email_id))
                        )).click()

                    child_email_list = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located(
                            (By.XPATH,
                             '//*[@id="ConversationReadingPaneContainer"]/div[2]/div/div'))
                    )
                    try:
                        child_email_list.find_element(By.XPATH,
                                                      'div[1]/div/div/div/div/div[1]/div[4]/div[1]/div/div[3]/button').click()
                        WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH,
                                                        '//*[starts-with(@id, "docking_InitVisiblePart_")]/div/div[2]/div[1]/div/span/button[1]/span/i'))
                        ).click()
                        logging.info('Resend:\nid: {}\nbody:{}'.format(email_id, email_body))
                        processed_email.append(email_id)
                        logging.info('{} spam emails resent.'.format(str(len(processed_email))))
                    except NoSuchElementException:
                        processed_email.append(email_id)
                        logging.info(
                            "Skipped:\nid: {}\nbody:{}\nresend button does not exist.".format(email_id, email_body))
                else:
                    processed_email.append(email_id)
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//span[contains(@class, 'gtcPn') and text()='Inbox']"))
            ).click()
        except TimeoutException:
            logger.error("Timeout occurred while waiting for elements.")
        except Exception as e:
            logger.error(f"An error occurred: {str(e)}")
        time.sleep(10)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('-e', '--email', dest='email', help='Prod email inbox address',required=True)
    parser.add_argument('-p', '--password', dest='password', help='Prod email inbox password',required=True)
    args = parser.parse_args()
    login_to_outlook(args.email, args.password)
    monitor_spam()
