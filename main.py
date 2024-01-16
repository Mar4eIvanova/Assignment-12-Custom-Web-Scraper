import csv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time
import smtplib
import pandas as pd
from openpyxl.workbook import Workbook
from email.message import EmailMessage
import os

MY_EMAIL = os.environ.get("SENDER")
MY_PASSWORD = os.environ.get("PASSWORD")
RECIPIENT = os.environ.get("RECIPIENT")

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

url = "https://www.coindesk.com/data/"
path = "C:/Users/mivan/Downloads/chromedriver_win32"

driver.get(url)
driver.implicitly_wait(0.5)

namel = []
pricel = []
percentagel = []

# find coin name and store the value in list
coin = driver.find_elements(By.CSS_SELECTOR, "h2[class='inner-column']")

for c in coin[:10]:
    print(c.text.replace("\n", " "))
    namel.append(c.text.replace("\n", " "))
time.sleep(2)

# find price and store the value in list
price = driver.find_elements(By.CSS_SELECTOR, "span[class='typography__StyledTypography-sc-owin6q-0 iAXWoh']")

for p in price[:10]:
    print(p.text)
    pricel.append(p.text)
time.sleep(2)

# find percentage and store the value in list
percent = driver.find_elements(By.CLASS_NAME, "percentage")

for perc in percent[:10]:
    print(perc.text)
    percentagel.append(perc.text)
time.sleep(2)

# combining the lists
data = []

for i in range(0, 10):
    data.append(namel[i])
    data.append(pricel[i])
    data.append(percentagel[i])
print(data)

# Split into small list of length 3.
final_data = [data[i:i + 3] for i in range(0, len(data), 3)]
print(final_data)

# write csv file
with open('10 Top Daily Coin Price.csv', 'w', newline='') as file:
    writer = csv.writer(file, dialect='excel')
    field = ["CoinName", "Price", "Percentage"]

    writer.writerow(field)
    writer.writerows(final_data)

new_dataFrame = pd.read_csv('10 Top Daily Coin Price.csv')
new_excel = pd.ExcelWriter('10 Top Daily Coin Price.xlsx')
new_dataFrame.to_excel(new_excel, index=False)
new_excel.close()


def send_mail_with_excel(recipient_email, subject, content, excel_file):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = MY_EMAIL
    msg['To'] = recipient_email
    msg.set_content(content)

    with open(excel_file, 'rb') as f:
        file_data = f.read()
    msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename=excel_file)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(MY_EMAIL, MY_PASSWORD)
        smtp.send_message(msg)


send_mail_with_excel(RECIPIENT,
                     "Daily Crypto Price",
                     "Here is daily price for top 10 market cap crypto currencies ",
                     "10 Top Daily Coin Price.xlsx")

driver.quit()


