#!  /bin/python
#   
#   Author: Walker Price
#   Description: This script is used to check the toner levels of all the printers on premesis.
#   Usage: python tonerck.py
#   Dependencies:
#       - Selenium
#       - printers.csv inventory file
#   To install dependencies run: pip install -r requirements.txt

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import csv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import os

opt = Options()
opt.add_argument("--headless")
driver = webdriver.Chrome(options=opt)
# Uncomment following line for debug purposes
# driver = webdriver.Chrome()

def main():
    toner_levels = []
    with open('printers.csv', 'r') as inventory:
        printers = csv.DictReader(inventory)
        for printer in printers:
            if "Toshiba" not in printer['Model']:
                toner_level = get_toner_levels(printer['IP'])
                print(f"{printer['Name']}: {toner_level * 100}%")
                toner_levels.append({
                    'Name': printer['Name'], 
                    'Location': printer['Location'], 
                    'Model': printer['Model'], 
                    'IP': printer['IP'],
                    'Black Toner': printer['Black Toner'],
                    'Color Toner': printer['Color Toner'],
                    'Toner Level': f"{toner_level * 100}%"
                })

    # Write results to toner_levels.csv
    with open('toner_levels.csv', 'w', newline='') as tonerfile:
        writer = csv.DictWriter(tonerfile, fieldnames=['Name', 'Location', 'Model', 'IP', 'Black Toner', 'Color Toner', 'Toner Level'])
        for result in sorted(toner_levels, key=lambda e: float(e['Toner Level'][:-1])):
            writer.writerow(result)
    driver.close()

    # Send email to wig_itsupport@delawarenorth.com
    send_email(toner_levels)


def get_toner_levels(printer_ip):
    driver.get(f"http://{printer_ip}/web/guest/en/websys/webArch/getStatus.cgi")
    kcmy = driver.find_elements(by=By.CLASS_NAME, value="bdr-1px-666")
    k = int(kcmy[0].get_attribute("width"))
    return k / 160


def send_email(toner_levels):
    sender = os.getenv('TONERCK_SENDER')
    recipient = os.getenv('TONERCK_RECIPIENT')
    smtp = os.getenv('TONERCK_SMTP')

    # Construct Message
    message = MIMEMultipart("alternative")
    message["Subject"] = "Automated Toner Check"
    message["From"] = sender
    message["To"] = recipient

    html = "<html><body><ul style='list-style: none'>"
    html += f"<p>Toner Check ran at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>"
    for printer in sorted(toner_levels, key=lambda e: float(e['Toner Level'][:-1])):
        lvl = float(printer['Toner Level'][:-1])
        textcolor = 'red' if lvl <= 20 else ('yellow' if lvl <= 50 else '#00FFD0')
        html += f"<li style='color: {textcolor}'>{printer['Name']} (Cartridge {printer['Black Toner']}): {printer['Toner Level']}</li>"
    html += '</ul></body></html>'

    # Set content type and attach content to message
    part1 = MIMEText(html, "html")
    message.attach(part1)

    # Send
    s = smtplib.SMTP(smtp)
    s.sendmail(sender, recipient, message.as_string())
    s.quit()

if __name__ == "__main__":
    main()