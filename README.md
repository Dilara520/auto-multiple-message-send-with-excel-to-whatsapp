# WhatsApp Message Sender with Selenium

This project automates the process of sending messages through WhatsApp Web using Selenium. It reads contact information and message details from an Excel file and sends personalized messages to each contact. Additionally, it can attach and send media files.

## Prerequisites

Before you start, ensure you have the following installed on your system:

- Python 3.x
- Selenium
- Pandas
- Chrome browser
- ChromeDriver (compatible with your Chrome version)

## Installation

1. **Clone the repository:**

    ```sh
    git clone https://github.com/Dilara520/whatsapp-auto-send-excel.git 
    cd whatsapp-auto-send-excel
    ```

3. **Download ChromeDriver:**

    - Download ChromeDriver from [here](https://sites.google.com/chromium.org/driver/downloads).
    - Ensure it is compatible with your Chrome version.
    - Place it in a directory included in your system's PATH.

## Usage

1. **Prepare your Excel file:**

    - Ensure your Excel file contains the following columns:
        - `Telefon`: The phone number of the contact.
        - `Açıklama 1`: First part of the message. (example: 'Merhabalar')
        - `Açıklama 2`: Second part of the message. (example: 'Name Surname,')
        - `Mesaj`: The main message content. (`%0A`: new line)
        - `Medya`: (Optional) The path to the media file to be sent along with the message.

2. **Update the script:**

    - Open `wpselen.py` and update the `file_path` variable to the path of your Excel file.
    - Update the `chrome_user_data_dir` variable with the path to your Chrome user data directory.
        - On macOS: `'/Users/<YOUR_USER_NAME>/Library/Application Support/Google/Chrome'`
        - On Windows: `'C:\\Users\\<YOUR_USER_NAME>\\AppData\\Local\\Google\\Chrome\\User Data'`

3. **Run the script:**

    ```sh
    python wpselen.py
    ```

    - The script will open a Chrome browser window and navigate to WhatsApp Web.
    - You will need to scan the QR code to log in (if not already logged in).
    - Once logged in, press `Enter` to continue the script.
    - The script will read the Excel file and send the messages to the specified contacts.

## Notes

- **QR Code Login:** The script uses your existing Chrome user data directory to avoid repeated QR code scans. Ensure you have proper permissions and the path is correctly set.
- **Media Files:** If a media path is provided, the script will attach and send the media file along with the message.

## Troubleshooting

- **QR Code Login Issue:** If the script opens a new Chrome window and requires a QR code scan each time even you are already logged in, ensure the `chrome_user_data_dir` is correctly set and accessible.
- **Permissions:** Ensure the Chrome user data directory has the necessary read and write permissions for the user running the script.
- **ChromeDriver:** Ensure ChromeDriver is compatible with your installed Chrome version and is in your system's PATH.

