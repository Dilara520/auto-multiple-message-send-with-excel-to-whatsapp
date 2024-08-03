import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
import time

#file_path = '<your_path_to_the_file>'
#chrome_user_data_dir_mac = '/Users/<your_username>/Library/Application Support/Google/Chrome'
#chrome_user_data_dir_win = r'C:\Users\<your_username>\AppData\Local\Google\Chrome\User Data'

df = pd.read_excel(file_path)

options = webdriver.ChromeOptions()
options.add_argument(f"user-data-dir={chrome_user_data_dir_win}") #chrome_user_data_dir_mac
options.add_argument("profile-directory=Default")

driver = webdriver.Chrome(options=options)

driver.get('https://web.whatsapp.com')
input("Hazır olunca Enter tuşuna basınız.")

def send_message(row, row_index):
    phone = row['Telefon']
    msg1 = row['Açıklama 1']
    msg2 = row['Açıklama 2']
    msg3 = row['Mesaj']
    media_path = row['Medya']
    
    message_parts = [msg1, msg2, msg3]
    message = " ".join(part for part in message_parts if pd.notna(part) and part not in [None, ''])
    
    try:
        new_chat_button = WebDriverWait(driver, 10).until( #//*[@id="app"]/div/div[2]/div[3]/header/header/div/span/div/span/div[1]/div/span
            EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div/div[2]/div[3]/header/div[2]/div/span/div[4]/div/span'))
        )
        new_chat_button.click()

        search_bar = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div[1]/span/div/span/div/div[1]/div[2]/div[2]/div/div[1]/p'))
        )
        search_bar.send_keys(phone)
        
        time.sleep(1)

        try:
            contact = driver.find_element(By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div[1]/span/div/span/div/div[2]/div/div/div/div[2]') #contact
        except:
            try:
                contact = driver.find_element(By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div[1]/span/div/span/div/div[2]/div[2]') #unknown
            except:
                # If no contact is found, move to the next row
                return_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div[1]/span/div/span/div/header/div/div[1]/div/span'))
                )
                return_button.click()
                return
        contact.click()

        time.sleep(1)
        message_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p'))
        )
        # Send the message line by line, using SHIFT + ENTER for new lines
        for line in message.split('\n'):
            message_input.send_keys(line)
            message_input.send_keys(Keys.SHIFT + Keys.ENTER)
        
        # Remove the last SHIFT + ENTER
        message_input.send_keys(Keys.BACKSPACE)
        time.sleep(1)

        if pd.notna(media_path) and media_path not in [None, '']:
            attach_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "#main > footer > div._ak1k.xnpuxes.copyable-area > div > span:nth-child(2) > div > div._ak1t._ak1m > div._ak1o > div > div"))
            )
            attach_button.click()

            image_box = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#main > footer > div._ak1k.xnpuxes.copyable-area > div > span:nth-child(2) > div > div._ak1t._ak1m > div._ak1o > div > span > div > ul > div > div:nth-child(2) > li > div > input"))
            )
            image_box.send_keys(media_path)

            send_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div/span'))
            )
            send_button.click()
        
        else:
            send_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]'))
            )
            send_button.click()

        time.sleep(2)
    except Exception as e:
        print(f"Failed to send message to {phone}. Error: {e}")

for index, row in df.iterrows():
    send_message(row, index)

driver.quit()