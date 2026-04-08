import pandas as pd
import time
import io
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from docx import Document
from docx.shared import Pt, Inches
import matplotlib.pyplot as plt

# --- NASTAVITVE ---
USER_NAME = os.environ.get('SLOFIT_USERNAME')
PASSWORD = os.environ.get('SLOFIT_PASS')
if not PASSWORD:
    raise ValueError("Napaka: Okoljska spremenljivka 'SLOFIT_PASS' ni nastavljena!")

# --- PREBERI EXCEL ---
file_path = "reportOKSanaliza.xlsx"
df_analize = pd.read_excel(file_path, sheet_name="Analize")
df_centili = pd.read_excel(file_path, sheet_name="Centili")

# --- VNOS ANALIZE ---
analiza_id = input("Vnesi številko analize (Meritev.id): ")
analiza_id = int(analiza_id)
rezultat = df_analize[df_analize["Meritev.id"] == analiza_id]

if not rezultat.empty:
    vrstica = rezultat.iloc[0]
    ID = vrstica["ID"]
    Meritev_id = vrstica["Meritev.id"]
    Regija_id = vrstica["Regija.id"]
    ReportDocName = vrstica["ReportDocName"]
    Opomba = vrstica["Opomba"]
else:
    raise ValueError("Analiza ni bila najdena.")

# --- SQL POIZVEDBA 1 ---
SQL_QUERY_1 = f"""
DECLARE @merID INT  = {Meritev_id}
DECLARE @regijaID INT = {Regija_id}

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

# --- SQL POIZVEDBA 2 ---
SQL_QUERY_2 = f"""
DECLARE @merID INT = {Meritev_id};

SELECT
    te.Ime AS TestIme,
    AVG(r.CentilniRang1) AS Povprecje
FROM Merjenec m
JOIN Clan c ON c.ID = m.IdClan
JOIN Rezultat r ON r.IdMerjenec = m.ID
JOIN Test te ON te.ID = r.IdTest
WHERE r.CentilniRang1 IS NOT NULL
AND EXISTS (
    SELECT 1
    FROM Meritev me2
    JOIN Merjenec m2 ON me2.ID = m2.IdMeritev
    WHERE m2.IdClan = m.IdClan
      AND me2.BaterijaTestov = 26
)
AND m.IdMeritev = @merID
GROUP BY te.Ime
ORDER BY te.Ime;
"""

# --- SELENIUM PRIJAVA ---
chrome_options = Options()
# chrome_options.add_argument("--headless")  # Odkomentiraj za headless
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)
wait = WebDriverWait(driver, 20)

def run_sql_and_get_table(sql_query):
    """Funkcija za izvedbo SQL poizvedbe preko spletnega UI in vračanje Pandas DataFrame."""
    driver.get("https://moj.slofit.org/Moj-SLOfit/sql")
    wait.until(EC.presence_of_element_located((By.CLASS_NAME, "CodeMirror")))
    js_query = sql_query.replace("\\", "\\\\").replace("\n", "\\n")
    driver.execute_script(f"""
    var editor = document.querySelector('.CodeMirror').CodeMirror;
    editor.setValue('{js_query}');
    editor.save();
    """)
    time.sleep(1)
    run_button = wait.until(EC.presence_of_element_located((By.ID, "dnn_ctr1155_SQL_cmdExecute")))
    driver.execute_script("arguments[0].click();", run_button)
    time.sleep(2)
    html_buffer = io.StringIO(driver.page_source)
    tables = pd.read_html(html_buffer)
    if tables:
        df_res = max(tables, key=len)
        if "Next" in str(df_res.iloc[-1].values):
            df_res = df_res.iloc[:-1]
        return df_res
    else:
        return pd.DataFrame()

try:
    # --- Prijava ---
    driver.get("https://moj.slofit.org/")

    # Vpis uporabniškega imena
    user_field = wait.until(EC.presence_of_element_located(
        (By.ID, "dnn_ctr543_View_ctlWrapper1_ctlLogin1_ctlUserName_txtTextBox")))
    user_field.clear()
    user_field.send_keys(USER_NAME)

    # Vpis gesla
    pass_field = wait.until(EC.presence_of_element_located(
        (By.ID, "dnn_ctr543_View_ctlWrapper1_ctlLogin1_ctlPassword_txtTextBox")))
    pass_field.clear()
    pass_field.send_keys(PASSWORD)

    # Klik na gumb PRIJAVA
    login_button = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//a[contains(text(),'PRIJAVA')] | //input[@value='PRIJAVA']")))
    driver.execute_script("arguments[0].click();", login_button)

    # Čakamo na pozdravno sporočilo
    wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Pozdravljeni!')]")))
    
    # --- SQL 1 ---
    df_sql1 = run_sql_and_get_table(SQL_QUERY_1)
    if df_sql1.empty:
        raise ValueError("SQL poizvedba 1 ni vrnila podatkov.")
    row = df_sql1.iloc[0]

    # --- SQL 2 ---
    df_sql2 = run_sql_and_get_table(SQL_QUERY_2)
    if df_sql2.empty:
        raise ValueError("SQL poizvedba 2 ni vrnila podatkov.")

finally:
    driver.quit()

# --- PRIPRAVA TABELE ZA DOCX ---


def place_decimal_and_round(number):
   
    
    num_str = str(number)
    
    
    
    
    new_num = float(num_str[:2] + '.' + num_str[2:])
    
    
    return round(new_num, 1)


# --- PRIPRAVA TABELE ---
df_table = pd.DataFrame()
df_table["Test"] = df_sql2["TestIme"]
df_sql2["Povprecje"] = df_sql2["Povprecje"].apply(place_decimal_and_round)
df_table["ZmigajSdv"] = df_sql2["Povprecje"]

df_table["Osrednjeslovenska"] = df_centili[df_centili["Geo_enota"] == "Osrednjeslovenska"]["CentilXA"].values[:len(df_table)]

df_table["Slovenija"] = df_centili[df_centili["Geo_enota"] == "Slovenija"]["CentilXA"].values[:len(df_table)]

# --- USTVARJANJE DOCX ---
doc = Document()
doc.add_heading('Analitično poročilo za meritev Zmigaj se do vadbe', level=1)
doc.add_heading('Osnovni podatki o meritvi', level=2)
doc.add_paragraph("")
doc.add_paragraph(f"Datum meritev: {row['Datum']}")
doc.add_paragraph(f"Izvajalec meritev: {row['IzvajalecNaziv']}")
doc.add_paragraph(f"Administrator: {row['Administrator']}")
doc.add_paragraph(f"Testna baterija: {row['BaterijaTestovIme']}")
doc.add_paragraph(f"Regija (za primerjavo): {row['RegijaName']}")
doc.add_paragraph(f"Število udeležencev ZV: {row['Merjencev']}")
doc.add_paragraph("")

# --- TABELA ---
doc.add_heading('Povprečni centili po testih', level=2)
table = doc.add_table(rows=1, cols=len(df_table.columns))
table.style = 'Table Grid'

# Glava
hdr_cells = table.rows[0].cells
for i, col_name in enumerate(df_table.columns):
    hdr_cells[i].text = col_name

# Vrstice
for idx, r in df_table.iterrows():
    row_cells = table.add_row().cells
    for i, col_name in enumerate(df_table.columns):
        row_cells[i].text = str(round(r[col_name], 2)) if isinstance(r[col_name], (int,float)) else str(r[col_name])

doc.add_paragraph("")

# --- GRAF ---
plt.figure(figsize=(8,4))

x = df_table["Test"]

y1 = df_table["ZmigajSdv"] - 50
y2 = df_table["Osrednjeslovenska"] - 50
y3 = df_table["Slovenija"] - 50

positions = list(range(len(x)))

plt.bar([p-0.2 for p in positions], y1, width=0.2, label="ZmigajSdv")
plt.bar(positions, y2, width=0.2, label="Osrednjeslovenska")
plt.bar([p+0.2 for p in positions], y3, width=0.2, label="Slovenija")

# OSI
plt.ylabel("Povprečni centil")
plt.ylim(-50, 50)
plt.yticks([-50, -25, 0, 25, 50], [0, 25, 50, 75, 100])

# sredina = 50
plt.axhline(0, linestyle='--')

plt.xticks(positions, x, rotation=45, ha='right')
plt.legend(loc='lower center', bbox_to_anchor=(0.5, 1.02), ncol=3)
plt.tight_layout()


graf_path = os.path.join(os.getcwd(), "graf.png")
plt.savefig(graf_path)
plt.close()
doc.add_picture(graf_path, width=Inches(6))

# --- SHRANJEVALNJE DOCX ---
docx_filename = os.path.join(os.getcwd(), ReportDocName + ".docx")
doc.save(docx_filename)
print(f"DOCX poročilo shranjeno v: {docx_filename}")