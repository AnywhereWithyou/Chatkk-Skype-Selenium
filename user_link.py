import time
import pandas as pd
import keyboard

from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

from restricted_input import r_input

from setting import *
from config import *
from functions import *
import tkinter as tk

    
profile_id = ''

# Delete Profile
try:
    profile_id = fnGetUUID(f'{OCTO_ID}')
    stop_profile(profile_id)
    deleteProfile(profile_id)
    print(f'Success to delete {OCTO_ID} profile!')
except:
    print(f'There does not exist with profile name {OCTO_ID}')

# # Create Profile
try:
    profile_id = fnGetUUID(f'{OCTO_ID}')
except:
    print(f'Create Octo Profile with {OCTO_ID}.')
    profile_id = createProfile(f'{OCTO_ID}')


port = get_debug_port(profile_id)
driver = get_webdriver(port)
links_list = []
stop_script = False
def stop_program(e):
    global stop_script
    if e.name == "home":
        stop_script = True

# Register the Home key event listener
keyboard.on_press(stop_program)

try:
    for i in range(1, 1001):
        if stop_script:
            break  # Exit the loop if the Home key is pressed
        driver.get(CHATKK_URL + f"/{i}")
        time.sleep(1)

        # Wait for the list to become visible (adjust the timeout as needed)
        ul_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'x5ul')))

        # Find all the <li> elements within the <ul> element
        li_elements = ul_element.find_elements(By.XPATH, './/li')

        # Check if li_elements is empty
        if not li_elements:
            print(f"No elements found for {i}, breaking loop.")
            break

        # Extract and print the links from each <li> element
        for li in li_elements:
            a_element = li.find_element(By.XPATH, './/a[@class="x5a0"]')
            link = a_element.get_attribute('href')
            print(link)
            links_list.append(link)

finally:
    # Close the WebDriver
    driver.quit()

# Create a DataFrame from the links_list
links_df = pd.DataFrame({'Links': links_list})

# Save the DataFrame to an Excel file
links_df.to_excel('links.xlsx', index=False)