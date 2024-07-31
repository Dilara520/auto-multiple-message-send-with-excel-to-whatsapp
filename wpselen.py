import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
import time

file_path = '<PATH_TO_YOUR_EXCEL_FILE>'
df = pd.read_excel(file_path)

#chrome_user_data_dir_mac = '/Users/<YOUR_USER_NAME>/Library/Application Support/Google/Chrome'
#chrome_user_data_dir_win = 'C:\Users\<YOUR_USER_NAME>\AppData\Local\Google\Chrome\User Data'

options = webdriver.ChromeOptions()
options.add_argument(f"user-data-dir={chrome_user_data_dir_mac}")

driver = webdriver.Chrome(options=options)

driver.get('https://web.whatsapp.com')
input("Whatsapp QR Kodunu Okutup Giriş Yap, Enter Tuşuna Bas.")

for index, row in df.iterrows():
    phone = row['Telefon']
    msg1 = row['Açıklama 1']
    msg2 = row['Açıklama 2']
    msg3 = row['Mesaj']
    media_path = row['Medya']
    
    message = f" {msg1} {msg2} {msg3}"

    url = f"https://web.whatsapp.com/send?phone={phone}&text={message}"
    driver.get(url)
    
    driver.implicitly_wait(10)

    try:
        if pd.notna(media_path) and media_path not in [None, '']:
            attach_button = driver.find_element(By.CSS_SELECTOR, "#main > footer > div._ak1k.xnpuxes.copyable-area > div > span:nth-child(2) > div > div._ak1t._ak1m > div._ak1o > div > div")
            attach_button.click()

            # Find and upload the image
            image_box = driver.find_element(By.CSS_SELECTOR, "#main > footer > div._ak1k.xnpuxes.copyable-area > div > span:nth-child(2) > div > div._ak1t._ak1m > div._ak1o > div > span > div > ul > div > div:nth-child(2) > li > div > input")
            image_box.send_keys(media_path)
            time.sleep(3)  # Wait for the image to load

            send_button = driver.find_element(By.CSS_SELECTOR, "#app > div > div.two._aigs > div._aigu > div._aigv._aigz > span > div > div > div > div.x1n2onr6.xyw6214.x78zum5.x1r8uery.x1iyjqo2.xdt5ytf.x1hc1fzr.x6ikm8r.x10wlt62.x1tkvqr7 > div > div.x78zum5.x1c4vz4f.x2lah0s.x1helyrv.x6s0dn4.x1qughib.x178xt8z.x13fuv20.x1nfbk4f.x1y1aw1k.xwib8y2.x1d52u69.xktsk01 > div.x1247r65.xng8ra > div > div > span")
            send_button.click()
        
        else:
            send_button = driver.find_element(By.XPATH, '//span[@data-icon="send"]')
            send_button.click()

        time.sleep(5)  # wait before moving to the next contact

    except Exception as e:
        print(f"Failed to send message to {phone}. Error: {e}")

driver.quit()