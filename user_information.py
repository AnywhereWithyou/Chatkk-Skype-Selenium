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

stop_script = False
def stop_program(e):
    global stop_script
    if e.name == "home":
        stop_script = True

# Register the Home key event listener
keyboard.on_press(stop_program)
# Read the Excel file containing URLs

# Read the Excel file containing URLs
df = pd.read_excel('links.xlsx')

# Create empty lists to store data
userURL_list = []
name_list = []
gender_list = []
country_list = []
birthday_list = []
age_list = []
skype_id_list = []

try:
    # Iterate through each row in the DataFrame
    index = 115
    for index, row in df.iterrows():
        if stop_script:
            break  # Exit the loop if the Home key is pressed
        if index == 500:break

        userURL = row['Links']  # Assuming 'userURL' is the column name in your Excel file

        driver.get(userURL)
        time.sleep(1)

        # Wait for the div with id "zd2" to become visible
        div_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'zd2')))

        # Initialize variables to store user information
        name = ""
        gender = ""
        country = ""
        birthday = ""
        age = ""
        skype_id = ""

        # Find all the <li> elements within the <ul> elements in the "zd2" div
        li_elements = div_element.find_elements(By.XPATH, './/ul[@class="zul"]/li')

        # Loop through the li_elements to extract and store user information
        for li in li_elements:
            span1 = li.find_element(By.CLASS_NAME, 'zs1')
            span2 = li.find_element(By.CLASS_NAME, 'zs')
            key = span1.text.strip()
            value = span2.text.strip()

            # Check the key and update the corresponding variable
            if key == "Name":
                name = value
            elif key == "Gender":
                gender = value
            elif key == "Country":
                country = value
            elif key == "Birthday":
                birthday = value
            elif key == "Age":
                age = value
            elif key == "Skype ID":
                skype_id = value

        # Append the data to the respective lists
        userURL_list.append(userURL)
        name_list.append(name)
        gender_list.append(gender)
        country_list.append(country)
        birthday_list.append(birthday)
        age_list.append(age)
        skype_id_list.append(skype_id)

                # Print the user's information
        print(f"User Information for {userURL}:")
        print(f"Name: {name}")
        print(f"Gender: {gender}")
        print(f"Country: {country}")
        print(f"Birthday: {birthday}")
        print(f"Age: {age}")
        print(f"Skype ID: {skype_id}")
        print("-" * 50)  # Separator

except Exception as e:
    print(f"Error: {str(e)}")

finally:
    # Create a new DataFrame from the lists
    data = {
        "userURL": userURL_list,
        "Name": name_list,
        "Gender": gender_list,
        "Country": country_list,
        "Birthday": birthday_list,
        "Age": age_list,
        "Skype ID": skype_id_list
    }
    new_df = pd.DataFrame(data)

    # Save the new DataFrame to a CSV file
    new_df.to_excel('user_information_115_500.xlsx', index=False)

    # Close the WebDriver
    driver.quit()