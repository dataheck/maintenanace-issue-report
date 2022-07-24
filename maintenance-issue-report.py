from dotenv import load_dotenv
from github import Github
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep
import chromedriver_binary # pylint: disable=unused-import
import json
import os
import selenium.webdriver.support.ui as ui
import tempfile
from functools import reduce

PRINT_DIALOG_DELAY = 3000 # milliseconds
PRINT_SAVE_DELAY = 500 # milliseconds
INTERACTIVE_TIMEOUT = 320 # seconds

load_dotenv()

mandatory_configuration = ('GITHUB_API_KEY', 'GITHUB_ORGANIZATION','GITHUB_PROJECT_NAME', 'GITHUB_PROJECT_FINISHED_COLUMN', 'PDF_SAVE_PATH')
config = dict()

for key in mandatory_configuration:
    value = os.environ.get(key)
    assert value is not None, f"Please set {key} before proceeding."
    config[key] = value

g = Github(config['GITHUB_API_KEY'])
org = g.get_organization(config['GITHUB_ORGANIZATION'])
project = None

for this_project in org.get_projects():
    if this_project.name == config['GITHUB_PROJECT_NAME']:
        project = this_project
        break

assert project is not None, f"Project '{config['GITHUB_PROJECT_NAME']}' was not found in the organization, please verify configuration."

project_column = None

for this_column in project.get_columns():
    if this_column.name == config['GITHUB_PROJECT_FINISHED_COLUMN']:
        project_column = this_column
        break

assert project_column is not None, f"Project column '{config['GITHUB_PROJECT_FINISHED_COLUMN']}' was not found in the project, please verify configuration."

# https://stackoverflow.com/questions/47007720/set-selenium-chromedriver-userpreferences-to-save-as-pdf
settings = {
    "recentDestinations": [{
        "id": "Save as PDF",
        "origin": "local",
        "account": "",
    }],
    "selectedDestinationId": "Save as PDF",
    "version": 2
}

selection_rules = {
    "kind": "local",
    "idPattern": "*",
    "namePattern": "Save as PDF"
}

prefs = {
    'printing.print_preview_sticky_settings.appState': json.dumps(settings), 
    'printing.default_destination_selection_rules': json.dumps(selection_rules),
    'savefile.default_directory': config['PDF_SAVE_PATH']
}

chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option('prefs', prefs)
chrome_options.add_argument('--kiosk-printing')

driver = webdriver.Chrome(options=chrome_options)

driver.get("https://github.com/login")
print("Please login to GitHub in the open window.")

class LoginTagHasValue(object):
    """ Validates that there is a meta tag with the name 'user-login' that also has non-zero-length content. """
    def __init__(self, meta_name='user-login'):
        self.meta_name = meta_name
    
    def __call__(self, driver):
        element = driver.find_element(By.XPATH, f"//meta[@name='{self.meta_name}']")
        if len(element.get_attribute("content")) > 0:
            return element
        else:
            return False

wait = ui.WebDriverWait(driver, timeout=INTERACTIVE_TIMEOUT)
wait.until(LoginTagHasValue())

#action = ActionChains(driver)

for card in project_column.get_cards():
    this_issue = card.get_content()
    driver.get(this_issue.html_url)

    try:
        element = driver.find_element(By.CLASS_NAME, "markdown-title")
    except NoSuchElementException:
        print("Title at target URL doesn't match; we might not be logged in.")
        break

    #action.key_down(Keys.LEFT_CONTROL).send_keys("p").key_up(Keys.LEFT_CONTROL).perform() # trigger print dialog
    driver.execute_script("window.print();")

driver.close()