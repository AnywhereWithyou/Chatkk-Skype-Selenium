# from skpy import Skype
# sk = Skype("developerpassion12@gmail.com", "alwaysFriend.0514") # connect to Skype
# userID = ""
# userName = ""
# chatContent = f"How are you, {userName}"
# ch = sk.contacts[userID].chat

# ch.sendMsg(chatContent)

import pandas as pd
import os
from skpy import Skype
import keyboard
import time

# Function to get the last processed index from the text file
def get_last_processed_index():
    if os.path.exists('index.txt'):
        with open('index.txt', 'r') as index_file:
            last_processed_index = index_file.read().strip()
            if last_processed_index.isdigit():
                return int(last_processed_index)
    return None

# Function to save the current index to the text file
def save_last_processed_index(index):
    with open('index.txt', 'w') as index_file:
        index_file.write(str(index))

# Connect to Skype
sk = Skype("developerpassion12@gmail.com", "alwaysFriend.0514")

# Read the Excel file
df = pd.read_excel('user_information_115_500.xlsx')

# Get the last processed index from the text file
last_processed_index = get_last_processed_index()

# Start from the last processed index or 0 if the text file doesn't exist
start_index = last_processed_index if last_processed_index is not None else 0

stop_script = False
def stop_program(e):
    global stop_script
    if e.name == "home":
        stop_script = True

# Register the Home key event listener
keyboard.on_press(stop_program)

RichCountry = ["United State", "Canada"]
try:
    # Iterate through each row in the DataFrame, starting from the specified index
    for index, row in df.iterrows():
        if stop_script:
            break  # Exit the loop if the Home key is pressed
        if index <= start_index:
            continue  # Skip rows before the last processed index

        time.sleep(60)

        userID = (row['Skype ID']).lower()
        userName = row['Name']
        userCountry = row['Country']

        if userCountry in RichCountry:
            cost = "10%"  # 10% cost
        else:
            cost = "40$"  # Default cost

        # Check if userID and userName are not empty and not "Skype Private"
        if userID and userName and userID != "Shared Privately":
            if "live:" not in userID:
                userID = 'live:' + userID
            chatContent = f"How are you, {userName}\n" \
                        f"I came across your Skype contact information on the [Chatkk.com] website.\n" \
                        f"My name is Vito Coco, and I'm a programmer from Armenia. You might know buying and selling various accounts, such as game accounts, on online platforms[https://www.playerup.com/]. However, I'd like to clarify that my intention is different. I'm interested in the possibility of using your Upwork account to work as a developer on the platform.\n" \
                        f"Here's how it would work: On the 25th of each month, we would pay you {cost} from the earnings generated using your Upwork account. This could be a significant source of income for you, with minimal effort required on your part. Additionally, if you refer me to someone you know, there's a potential bonus for you.\n" \
                        f"I understand that your initial thought might be to question the legitimacy of this proposal. However, I want to emphasize that I have no ulterior motives or reasons to engage in any form of collaboration beyond this mutually beneficial arrangement. I believe in establishing a friendly and trustworthy partnership where we can both benefit.\n" \
                        f"If you find my proposal agreeable and are interested in discussing it further, please feel free to contact me via email or on Skype. \nMy email address is [developerpassion12@gmail.com].\n" \
                        f"Looking forward to hearing from you.\n" \
                        f"Best regards,\n" \
                        f"Vito Coco."
            print(chatContent)

            try:
                # Get the chat for the user
                ch = sk.contacts[userID].chat

                # Send the message
                ch.sendMsg(chatContent)

                # Print the message and index
                print(f"Message sent to {userName} (Skype ID: {userID}): {chatContent}")
                print(f"Processed up to index {index}")

            except Exception as e:
                # Print the error and index
                print(f"Failed to send message to {userName} (UserID: {userID}): {str(e)}")
                print(f"Processed up to index {index}")
                continue

except Exception as e:
    # Print any unexpected exceptions
    print(f"An unexpected error occurred: {str(e)}")

finally:
    # Save the current index to the text file
    save_last_processed_index(index)

