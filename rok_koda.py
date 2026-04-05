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
from docx import Document
from docx.shared import Pt

# --- NASTAVITVE ---
USER_NAME = os.environ.get('SLOFIT_USERNAME')  # Tukaj vpišite svoje podatke
PASSWORD = os.environ.get('SLOFIT_PASS')
if not PASSWORD:
    raise ValueError("Napaka: Okoljska spremenljivka 'SLOFIT_PASS' ni nastavljena!")
SQL_QUERY = f"""DECLARE @merID INT  = 47894 -- 47894;  -- Meritev.ID
DECLARE @regijaID INT = 8 -- 8;    -- FitnessMapRegions.ID
SELECT 
    m.Datum,
    im.Naziv AS IzvajalecNaziv,
    u.PriimekInIme AS Administrator,
    bt.Ime AS BaterijaTestovIme,
    fmr.Name AS RegijaName,
    (  SELECT COUNT(*)
        FROM Merjenec m1
        JOIN Clan c1 ON c1.ID = m1.IdClan
        WHERE m1.IdMeritev = @merID
          AND EXISTS (
                SELECT 1
                FROM Merjenec m2
                JOIN Meritev me2 ON me2.ID = m2.IdMeritev
                WHERE m2.IdClan = m1.IdClan
                  AND me2.BaterijaTestov = 26
          )
    ) AS Merjencev
FROM Meritev m
    INNER JOIN IzvajalecMeritev im 
        ON im.ID = m.Izvajalec
    INNER JOIN Uporabnik u 
        ON u.ID = m.IdUporabnik
    INNER JOIN BaterijaTestov bt
        ON bt.ID = m.BaterijaTestov
    INNER JOIN FitnessMapRegions fmr
        ON fmr.ID = @regijaID
WHERE m.ID = @merID;
"""

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
    
    poizvedba = SQL_QUERY
    
    try:
        # KLJUČNI POPRAVEK: Čakamo, da se CodeMirror element dejansko pojavi na strani
        # Čakamo do 20 sekund na element z razredom 'CodeMirror'
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "CodeMirror")))
        print("CodeMirror zaznan, vpisujem...")
        poizvedba_js = SQL_QUERY.replace("\\", "\\\\").replace("\n", "\\n")
        # Ta JS koda poišče CodeMirror instanco na strani in ji vnese tekst
        js_vpis = f"""
        var editor = document.querySelector('.CodeMirror').CodeMirror;
        editor.setValue('{poizvedba_js}');
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

# 8. Čakanje na tabelo in preverjanje dolžine (stranjenje)
    print("Čakam na izris rezultatov...")
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
    time.sleep(2)  # Kratek premor, da se naloži še DataTables vrstica

    try:
        # Preverimo, če je gumb 'Next' aktiven (nima razreda 'disabled')
        # XPath išče element, ki ima razred 'next', nima pa razreda 'disabled'
        next_button_xpath = "//a[contains(@class, 'next') and not(contains(@class, 'disabled'))]"
        
        is_paginated = len(driver.find_elements(By.XPATH, next_button_xpath)) > 0

        if is_paginated:
            print("Zaznanih več kot 50 vrstic. Preklapljam na prikaz 'All'...")
            
            # Najdemo dropdown za izbiro dolžine (iščemo po imenu, ki se konča na _length)
            from selenium.webdriver.support.ui import Select
            length_dropdown = driver.find_element(By.XPATH, "//select[contains(@name, '_length')]")
            
            select = Select(length_dropdown)
            # Izberemo možnost 'All', ki ima vrednost '-1'
            select.select_by_value("-1")
            
            print("Čakam, da se naložijo vsi podatki...")
            # Počakamo par sekund, da se tabela osveži z vsemi vrsticami
            time.sleep(4) 
        else:
            print("Vsi podatki so že vidni (manj kot 50 vrstic).")

    except Exception as e:
        print(f"Opomba pri preverjanju strani: {e}")
        # Nadaljujemo tudi če preklop spodleti, da vsaj shranimo kar je na voljo

    # Dejansko branje tabele v Pandas
    html_buffer = io.StringIO(driver.page_source)
    tables = pd.read_html(html_buffer)
    
    if tables:
        # Vedno izberemo tabelo z največ vrsticami
        df = max(tables, key=len) 
        
        # Odstranimo zadnjo vrstico, če Pandas po pomoti prebere 'Previous 1 2 Next' kot podatke
        if "Next" in str(df.iloc[-1].values):
            df = df.iloc[:-1]

        print(f"\nUspeh! Skupaj pridobljenih {len(df)} vrstic.")
        
        filename = "reportOKSanaliza.xlsx"
        df.to_excel(filename, index=False)
        print(f"Podatki so shranjeni v: {filename}")
        print("Ustvarjam DOCX poročilo...")

        # Preberemo Excel
        df_doc = pd.read_excel(filename)

        # Vzamemo drugo vrstico (index 0 = header, 1 = podatki)
        row = df_doc.iloc[0]

        # Ustvarimo dokument
        doc = Document()

        # Naslovi
        doc.add_heading('Uvodna stran', level=1)
        doc.add_heading('Analitično poročilo za meritev Zmigaj se do vadbe', level=2)
        doc.add_heading('Osnovni podatki o meritvi', level=3)

        # Prazna vrstica
        doc.add_paragraph("")

        # Podatki
        doc.add_paragraph(f"Datum meritev: {row['Datum']}")
        doc.add_paragraph(f"Izvajalec meritev: {row['IzvajalecNaziv']}")
        doc.add_paragraph(f"Administrator: {row['Administrator']}")
        doc.add_paragraph(f"Testna baterija: {row['BaterijaTestovIme']}")
        doc.add_paragraph(f"Regija (za primerjavo): {row['RegijaName']}")
        doc.add_paragraph(f"Število udeležencev ZV: {row['Merjencev']}")

        # Shrani DOCX v isto mapo
        docx_filename = os.path.join(os.getcwd(), "reportOKSanaliza.docx")
        doc.save(docx_filename)

        print(f"DOCX poročilo shranjeno v: {docx_filename}")
    else:
        print("Napaka: Tabela ni bila najdena.")


except Exception as e:
    print(f"\nPrišlo je do napake: {e}")

finally:
    print("\nPostopek zaključen. Brskalnik ostaja odprt za pregled.")
    # driver.quit() # Odkomentiraj za samodejno zapiranje