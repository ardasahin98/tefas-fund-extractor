import time
import os
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl

fonds = [
    "TI1", "TIV", "DCB", "BGP", "ALE", "TKM", "IGL", "APT",
    "AFT", "TTA", "IOG", "GO1", "GO3", "GO4", "TE3", "TPC"
]

# Create Excel
wb = openpyxl.Workbook()
ws = wb.active
ws.append(["Fund", "Price"])

options = Options()
options.add_argument("--disable-blink-features=AutomationControlled")
# IMPORTANT: do NOT use headless for TEFAS

service = Service(ChromeDriverManager().install())

for fond_name in fonds:
    driver = webdriver.Chrome(service=service, options=options)
    driver.get(f"https://www.tefas.gov.tr/FonAnaliz.aspx?FonKod={fond_name}")

    price = ""
    for _ in range(20):
        try:
            price = driver.find_element(
                By.XPATH,
                "//*[@id='MainContent_PanelInfo']//ul/li[1]/span"
            ).text
            if price.strip():
                break
        except:
            time.sleep(0.5)

    driver.quit()
    ws.append([fond_name, price])

# Save to Desktop with date
today = datetime.now().strftime("%Y-%m-%d")
desktop = os.path.join(os.path.expanduser("~"), "Desktop")
file_path = os.path.join(desktop, f"fund_values_{today}.xlsx")

wb.save(file_path)