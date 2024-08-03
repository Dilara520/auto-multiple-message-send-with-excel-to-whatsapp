import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
import time

file_path = '<PATH_TO_YOUR_EXCEL_FILE>'
df = pd.read_excel(file_path)

#chrome_user_data_dir_mac = '/Users/<YOUR_USER_NAME>/Library/Application Support/Google/Chrome'
#chrome_user_data_dir_win = 'C:\Users\<YOUR_USER_NAME>\AppData\Local\Google\Chrome\User Data'

options = webdriver.ChromeOptions()
options.add_argument(f"user-data-dir={chrome_user_data_dir_mac}") #chrome_user_data_dir_win
options.add_argument("profile-directory=Default")

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
    
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
    )

    try:
        if pd.notna(media_path) and media_path not in [None, '']:
            attach_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "#main > footer > div._ak1k.xnpuxes.copyable-area > div > span:nth-child(2) > div > div._ak1t._ak1m > div._ak1o > div > div"))
            )
            attach_button.click()

            # Wait for and upload the image
            image_box = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#main > footer > div._ak1k.xnpuxes.copyable-area > div > span:nth-child(2) > div > div._ak1t._ak1m > div._ak1o > div > span > div > ul > div > div:nth-child(2) > li > div > input"))
            )
            image_box.send_keys(media_path)

            # Wait for and click the send button
            send_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div/span'))
            )
            send_button.click()
        
        else:
            send_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]'))
            )
            send_button.click()

        # wait before moving to the next contact
        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element((By.XPATH, '//span[@data-icon="ptt"]'))
        )
        #update_excel(row_index, "Success")

    except Exception as e:
        print(f"Failed to send message to {phone}. Error: {e}")
        #update_excel(row_index, "Failed")

def update_excel(row_index, status):
    wb = load_workbook(file_path)
    ws = wb.active
    status_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    if status == "Success":
        status_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    elif status == "Failed":
        status_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    for cell in ws[row_index+2]:  # +2 to account for header row and 0-based index
        cell.fill = status_fill

    wb.save(file_path)
    wb.close()

for index, row in df.iterrows():
    send_message(row, index)

driver.quit()
