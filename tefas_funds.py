from flask import Flask, render_template, redirect, url_for
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl

app = Flask(__name__)

fonds = [
    "TI1", "TIV", "DCB", "BGP", "ALE", "TKM", "IGL", "APT",
    "AFT", "TTA", "IOG", "GO1", "GO3", "GO4", "TE3", "TPC"
]

def run_scraper():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Fund", "Price"])

    options = Options()
    options.add_argument("--disable-blink-features=AutomationControlled")
    # DO NOT use headless for TEFAS

    service = Service(ChromeDriverManager().install())

    for fond_name in fonds:
        driver = webdriver.Chrome(service=service, options=options)
        url = f"https://www.tefas.gov.tr/FonAnaliz.aspx?FonKod={fond_name}"
        driver.get(url)

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

    wb.save("tefas_prices.xlsx")

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/run")
def run():
    run_scraper()
    return redirect(url_for("index"))

if __name__ == "__main__":
    app.run(debug=True)