import selenium.webdriver as webdriver

from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import requests

from config import *

# Create Profile
def createProfile(profileName)->str:
    data = {
        "title": profileName,
        "fingerprint": {
                "os": "win",
                "screen": "1920x1080"
        },
        "proxy": PROXY
    }
    response = requests.post(url = OCTO_AUTOMATION_URL, json = data, headers=OCTO_HEADER).json()
    uuid = response.get('data')['uuid']
    return uuid

def deleteProfile(profileUUID)->str:
    data = {
        "uuids": [f"{profileUUID}"],
        "skip_trash_bin": True
    }
    response = requests.delete(url = OCTO_AUTOMATION_URL, json = data, headers=OCTO_HEADER).json()
    uuid = response.get('data')['deleted_uuids'][0]
    return uuid

# Get Profile uuid of OctoBrowser
def fnGetUUID(profileName):
    response_octo = requests.request("GET", f'{OCTO_AUTOMATION_URL}?search={profileName}', headers=OCTO_HEADER)
    data_uuid = response_octo.json()
    uuid = data_uuid.get('data')[0]['uuid']
    return uuid

# Get Debug Port
# Profile id is uuid from fnGetUUID()
def get_debug_port(profile_id):
    data = requests.post(
        f'{LOCAL_API}/start', json={'uuid': profile_id, 'headless': HEADLESS, 'debug_port': True}
    ).json()
    print(data)
    return data['debug_port']

def stop_profile(profile_id):
    data = requests.post(
        f'{LOCAL_API}/stop', json={'uuid': profile_id}
    ).json()
    print(data)

# Create webdriver
# port is from get_debug_port()
def get_webdriver(port):
    chrome_options = Options()
    chrome_options.add_experimental_option('debuggerAddress', f'127.0.0.1:{port}')
    # chrome_options.headless = True

    # chrome_options.add_argument("--headless=new")
    # chrome_options.add_experimental_option(
    #     'prefs', {'profile.default_content_setting_values.notifications': 1}
    # )
    # Change chrome driver path accordingly
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    return driver
