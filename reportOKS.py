import pandas as pd
import time
import io
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import os

# --- NASTAVITVE ---
USER_NAME = "bojan.leskosek@fsp.uni-lj.si"  # Tukaj vpišite svoje podatke
PASSWORD = os.environ.get('SLOFIT_PASS')
if not PASSWORD:
    raise ValueError("Napaka: Okoljska spremenljivka 'SLOFIT_PASS' ni nastavljena!")
SQL_QUERY = "SELECT TOP 3 * FROM clan"

chrome_options = Options()
# chrome_options.add_argument("--headless") # Zaženi brez vidnega okna (ko vse deluje)

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)
wait = WebDriverWait(driver, 20)

try:
    # 1. Odpri prijavno stran
    print("Odpiram prijavo na Moj-SLOfit...")
    driver.get("https://moj.slofit.org/")
    
    # 2. Vpis uporabniškega imena
    user_field = wait.until(EC.presence_of_element_located((By.ID, "dnn_ctr543_View_ctlWrapper1_ctlLogin1_ctlUserName_txtTextBox")))
    user_field.clear()
    user_field.send_keys(USER_NAME)
    
    # 3. Vpis gesla
    pass_field = driver.find_element(By.ID, "dnn_ctr543_View_ctlWrapper1_ctlLogin1_ctlPassword_txtTextBox")
    pass_field.clear()
    pass_field.send_keys(PASSWORD)
    
  

# 4. Klik na gumb za prijavo
    print("Iščem gumb PRIJAVA...")
    
    try:
        # Iščemo gumb, ki ima v besedilu 'PRIJAVA' (velike črke) 
        # ali 'Prijava' ali 'prijava' (uporabimo translate za "case-insensitive" iskanje)
        login_button = wait.until(EC.element_to_be_clickable((By.XPATH, 
            "//a[contains(translate(text(), 'prijava', 'PRIJAVA'), 'PRIJAVA')] | "
            "//input[contains(translate(@value, 'prijava', 'PRIJAVA'), 'PRIJAVA')] | "
            "//*[contains(@id, 'btnSubmit')]"
        )))
        
        # Uporabimo JavaScript klik, da preprečimo 'ElementClickIntercepted' napake
        driver.execute_script("arguments[0].click();", login_button)
        print("Gumb PRIJAVA uspešno pritisnjen.")
        
    except Exception as e:
        print(f"Gumba ni bilo mogoče klikniti: {e}")
        # Če zgornje ne dela, poskusiva še najbolj direkten XPath za ta gumb:
        login_button = driver.find_element(By.XPATH, "//a[@id='dnn_ctr543_View_ctlWrapper1_ctlLogin1_btnSubmit']")
        driver.execute_script("arguments[0].click();", login_button)


    
    # Počakamo, da se pojavi napis 'Pozdravljeni!', kar potrdi uspešno prijavo
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Pozdravljeni!')]")))
    print("Prijava uspešna.")

    # 5. Preklop na SQL stran
    driver.get("https://moj.slofit.org/Moj-SLOfit/sql")
    
# 6. Vnos poizvedbe neposredno v CodeMirror urejevalnik
    print("Vpisujem SQL poizvedbo v napredni urejevalnik...")
    
    poizvedba = "SELECT TOP 3 * FROM clan"
    
    try:
        # KLJUČNI POPRAVEK: Čakamo, da se CodeMirror element dejansko pojavi na strani
        # Čakamo do 20 sekund na element z razredom 'CodeMirror'
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "CodeMirror")))
        print("CodeMirror zaznan, vpisujem...")

        # Ta JS koda poišče CodeMirror instanco na strani in ji vnese tekst
        js_vpis = f"""
        var editor = document.querySelector('.CodeMirror').CodeMirror;
        editor.setValue('{poizvedba}');
        editor.save(); 
        """
        driver.execute_script(js_vpis)
        print("Vrednost vpisana v CodeMirror.")

    except Exception as e:
        print(f"CodeMirror ni bil najden ali pripravljen: {e}")
        print("Poskušam klasičen vpis v textarea...")
        
        # Rezervni načrt: Čakamo na textarea, če CodeMirrora slučajno ni
        try:
            sql_input = wait.until(EC.presence_of_element_located((By.NAME, "dnn$ctr1155$SQL$txtQuery")))
            driver.execute_script("arguments[0].value = arguments[1];", sql_input, poizvedba)
        except:
            print("Tudi klasično polje ni bilo najdeno. Preveri, če se je stran pravilno naložila.")
            raise  # Ponovno sproži napako, da vemo, kje se je ustavilo


    # 7. Klik na 'Run Script' 
    time.sleep(1)
    run_button = wait.until(EC.presence_of_element_located((By.ID, "dnn_ctr1155_SQL_cmdExecute")))
    driver.execute_script("arguments[0].click();", run_button)
    
    print("Gumb 'Run Script' pritisnjen.")

    print("Poizvedba poslana...")

    # 8. Čakanje na tabelo in branje
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
    time.sleep(3) # Dodaten čas za izris vsebine
    
    html_buffer = io.StringIO(driver.page_source)
    tables = pd.read_html(html_buffer)
    
    if tables:
        df = max(tables, key=len) # Izberi tabelo z največ podatki
        print(f"\nPridobljeno {len(df)} vrstic.")
        
        filename = "SLOfit_OKS_Izvoz.xlsx"
        #df.to_excel(filename, index=False)
        print(f"Podatki so shranjeni v: {filename}")
    else:
        print("Napaka: Tabela ni bila najdena.")

except Exception as e:
    print(f"\nPrišlo je do napake: {e}")

finally:
    print("\nPostopek zaključen. Brskalnik ostaja odprt za pregled.")
    # driver.quit() # Odkomentiraj za samodejno zapiranje