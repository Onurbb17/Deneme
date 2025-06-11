from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from datetime import datetime
import time

CHROMEDRIVER_PATH = "C:\\Users\\ONUR\\Desktop\\Yeni klasör (4)\\chromedriver.exe"  # Kendi yolunu güncel tut!

def akakce_en_uygun(linkler):
    urunler, fiyatlar, magazalar, links = [], [], [], []
    chrome_options = Options()
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("user-agent=Mozilla/5.0")
    service = Service(CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    for url in linkler:
        driver.get(url)
        # BEKLEME: Satıcılar yüklenene kadar bekle!
        try:
            wait = WebDriverWait(driver, 12)
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "ul#PL li")))
        except Exception:
            print("Satıcı listesi yüklenmedi, yine de deniyoruz...")
        try:
            urun = driver.find_element(By.TAG_NAME, "h1").text.strip()
        except Exception:
            urun = "Ürün adı bulunamadı"
        try:
            satici_listesi = driver.find_element(By.CSS_SELECTOR, "ul#PL")
            ilk_li = satici_listesi.find_elements(By.TAG_NAME, "li")[0]
            a_tag = ilk_li.find_element(By.CSS_SELECTOR, "a.iC.xt_v8")
            fiyat = a_tag.find_element(By.CSS_SELECTOR, ".pt_v8").text.strip().replace("\n", "").replace("\r", "")
            try:
                magaza = a_tag.find_element(By.CSS_SELECTOR, ".v_v8 b").text.strip()
            except Exception:
                magaza_text = a_tag.find_element(By.CSS_SELECTOR, ".v_v8").text.strip()
                magaza = magaza_text.split("/")[-1].strip()
        except Exception:
            fiyat = "Fiyat bulunamadı"
            magaza = "Mağaza bulunamadı"
        urunler.append(urun)
        fiyatlar.append(fiyat)
        magazalar.append(magaza)
        links.append(url)
        print(f"{urun}: {fiyat} - {magaza}")
    driver.quit()
    now = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    df = pd.DataFrame({
        "Ürün Adı": urunler,
        "Fiyat": fiyatlar,
        "Mağaza": magazalar,
        "Link": links
    })
    df.to_excel(f"akakce_en_uygun_{now}.xlsx", index=False)
    print("Excel kaydedildi.")

if __name__ == "__main__":
    linkler = [
        "https://www.akakce.com/scooterlar/en-ucuz-baby2go-racer-gri-isikli-3-tekerlekli-cocuk-fiyati,1043575320.html"
        # Diğer ürün linklerini de ekleyebilirsin
    ]
    akakce_en_uygun(linkler)