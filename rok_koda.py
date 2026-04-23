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
from docx.shared import Cm
import matplotlib.pyplot as plt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE

# --- NASTAVITVE ---
USER_NAME = os.environ.get('SLOFIT_USERNAME')
PASSWORD = os.environ.get('SLOFIT_PASS')
if not PASSWORD:
    raise ValueError("Napaka: Okoljska spremenljivka 'SLOFIT_PASS' ni nastavljena!")

# --- PREBERI EXCEL ---
file_path = "reportOKSanaliza.xlsx"
df_analize = pd.read_excel(file_path, sheet_name="Analize")
df_centili = pd.read_excel(file_path, sheet_name="Centili")
df_testi = pd.read_excel("reportOKSanaliza_2.xlsx", sheet_name="Testi")

# slovar: ime testa -> vrstni red
order_map = dict(zip(df_testi["Ime"], df_testi["Order"]))
visji_map = dict(zip(df_testi["Ime"], df_testi["VisjiBoljsi"]))

# --- VNOS ANALIZE ---
analiza_id = input("Vnesi ID analize: ")
analiza_id = int(analiza_id)

rezultat = df_analize[df_analize["ID"] == analiza_id]

if not rezultat.empty:
    vrstica = rezultat.iloc[0]
    table_style = vrstica["Table Style"]
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
    im.Naslov,
    u.PriimekInIme AS Administrator,
    bt.Ime AS BaterijaTestovIme,
    fmr.Name AS RegijaName,
    stats.Skupaj AS Merjencev,
    stats.Nm,
    stats.Nz,
    stats.XAstarost
FROM Meritev m
    INNER JOIN IzvajalecMeritev im ON im.ID = m.Izvajalec
    INNER JOIN Uporabnik u ON u.ID = m.IdUporabnik
    INNER JOIN BaterijaTestov bt ON bt.ID = m.BaterijaTestov
    INNER JOIN FitnessMapRegions fmr ON fmr.ID = @regijaID
    CROSS APPLY (
        SELECT
            COUNT(*) AS Skupaj,
            SUM(CASE WHEN c1.Spol = 1 THEN 1 ELSE 0 END) AS Nm,
            SUM(CASE WHEN c1.Spol = 2 THEN 1 ELSE 0 END) AS Nz,
            AVG(posameznik.StarostLet) AS XAstarost
        FROM Merjenec m1
        JOIN Clan c1 ON c1.ID = m1.IdClan
        CROSS APPLY (
            SELECT CAST(DATEDIFF(day, c1.DatumRojstva, m.Datum) / 365.25 AS FLOAT) AS StarostLet
        ) AS posameznik
        WHERE m1.IdMeritev = m.ID
    ) AS stats
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
SQL_QUERY_3 = f"""
DECLARE @merID INT = {Meritev_id};

SELECT
    te.Ime AS TestIme,
    SUM(CASE WHEN r.ConaBarva = '#9ac263' THEN 1 ELSE 0 END) AS Zdravo,
    SUM(CASE WHEN r.ConaBarva = '#ffd54e' THEN 1 ELSE 0 END) AS Mejno,
    SUM(CASE WHEN r.ConaBarva = '#df3127' THEN 1 ELSE 0 END) AS Tvegano
FROM Merjenec m
JOIN Clan c ON c.ID = m.IdClan
JOIN Rezultat r ON r.IdMerjenec = m.ID
JOIN Test te ON te.ID = r.IdTest
WHERE
    m.IdMeritev = @merID
    AND r.ConaBarva IN ('#9ac263', '#df3127', '#ffd54e')
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
    import json

    js_query = json.dumps(sql_query)

    driver.execute_script(f"""
    var editor = document.querySelector('.CodeMirror').CodeMirror;
    editor.setValue({js_query});
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
    kraj = row["Naslov"]
    st_moski = int(row["Nm"])
    st_zenske = int(row["Nz"])
    def fix_starost(val):
        s = str(val)

        # če je številka sumljivo dolga, popravi decimalno mesto
        if len(s) > 4:
            s = s[:2] + "." + s[2:]

        return round(float(s), 1)
    povp_starost = fix_starost(row["XAstarost"])

    # --- SQL 2 ---
    df_sql2 = run_sql_and_get_table(SQL_QUERY_2)
    if df_sql2.empty:
        raise ValueError("SQL poizvedba 2 ni vrnila podatkov.")
    # --- SQL 3 ----
    df_sql3 = run_sql_and_get_table(SQL_QUERY_3)

    if df_sql3.empty:
        raise ValueError("SQL poizvedba 3 ni vrnila podatkov.")

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
df_table["Order"] = df_table["Test"].map(order_map)
df_table = df_table.sort_values("Order")
df_table = df_table.drop(columns=["Order"])
def adjust_value(row, col):
    visji = visji_map.get(row["Test"], 1) 
    if visji == 0:
        return 100 - row[col]
    return row[col]

# popravi vse tri stolpce
for col in ["ZmigajSdv", "Osrednjeslovenska", "Slovenija"]:
    df_table[col] = df_table.apply(lambda r: adjust_value(r, col), axis=1)
def add_star(test):
    t = str(test).strip()
    if visji_map.get(t, 1) == 0:
        return t + "*"
    return t

df_table["Test"] = df_table["Test"].apply(add_star)

# --- USTVARJANJE DOCX ---
doc = Document("template.docx")
doc.add_heading('Analitično poročilo za meritev Zmigaj se do vadbe', level=1)

p = doc.add_paragraph("")  # prazen odstavek za razmik
p.paragraph_format.space_after = Pt(12)

doc.add_heading('Osnovni podatki o meritvi', level=2)
doc.add_paragraph("")
datum = pd.to_datetime(row["Datum"]).strftime("%d. %m. %Y")
doc.add_paragraph(f"Datum meritev: {datum}")
doc.add_paragraph(f"Izvajalec meritev: {row['IzvajalecNaziv']}")
doc.add_paragraph(f"Kraj izvajalca: {kraj}")
doc.add_paragraph(f"Administrator izvajalca: {row['Administrator']}")
doc.add_paragraph(f"Testna baterija: {row['BaterijaTestovIme']}")
doc.add_paragraph(f"Regija (za primerjavo): {row['RegijaName']}")
doc.add_paragraph(f"Število udeležencev skupine ZmigajSdv: {row['Merjencev']}")
doc.add_paragraph(f"Število moških: {st_moski}")
doc.add_paragraph(f"Število žensk: {st_zenske}")
doc.add_paragraph(f"Povprečna starost: {str(povp_starost).replace('.', ',')}")
doc.add_paragraph("")
doc.add_page_break()

# --- TABELA ---
doc.add_heading('Povprečni centili po testih', level=2)
doc.add_paragraph("Pri testih, pri katerih je zaželen nižji rezultat (npr. 6-minutni test hoje, Obseg pasu) so centili “obrnjeni” (100-centil), tako da v spodnji tabeli in grafikonu višje povprečje vedno pomeni boljši rezultat. Ti testi so označeni z *")
table = doc.add_table(rows=1, cols=len(df_table.columns))
if pd.notna(table_style) and str(table_style).strip() != "":
    try:
        table.style = str(table_style)
    except Exception:
        print(f"Opozorilo: stil '{table_style}' ne obstaja, uporabljen bo privzeti.")
        table.style = 'Table Grid'
else:
    table.style = 'Table Grid'
table.autofit = True
table.allow_autofit = True


# Glava
hdr_cells = table.rows[0].cells
for i, col_name in enumerate(df_table.columns):
    hdr_cells[i].text = col_name
    hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

# Vrstice
for idx, r in df_table.iterrows():
    row_cells = table.add_row().cells
    for i, col_name in enumerate(df_table.columns):
        val = str(round(r[col_name], 1)) if isinstance(r[col_name], (int, float)) else str(r[col_name])
        row_cells[i].text = val

        if i == 0:
            row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        else:
            row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
for row in table.rows:
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            paragraph.paragraph_format.keep_together = True
for row in table.rows:
    for cell in row.cells:
        cell.width = Inches(1.2)
table.columns[0].width = Cm(5.0)


for cell in table.columns[0].cells:
    cell.width = Cm(5.0)
doc.add_paragraph("")

# --- GRAF ---
plt.figure(figsize=(8,5))


x = df_table["Test"]

y1 = df_table["ZmigajSdv"] - 50
y2 = df_table["Osrednjeslovenska"] - 50
y3 = df_table["Slovenija"] - 50

positions = list(range(len(x)))

plt.bar([p-0.2 for p in positions], y1, width=0.2, label="ZmigajSdv", color='orange')
plt.bar(positions, y2, width=0.2, label="Osrednjeslovenska", color='#add8e6')  # svetlo modra
plt.bar([p+0.2 for p in positions], y3, width=0.2, label="Slovenija", color='#00008b')  # temno modra

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

# --- PREBERI REGIJO ---
df_regija = pd.read_excel(file_path, sheet_name="Regije")

regija_row = df_regija[df_regija["ID"] == Regija_id]

if regija_row.empty:
    raise ValueError("Regija ni bila najdena.")

regija_name = regija_row.iloc[0]["Name"]

# --- PREBERI CONE ---
df_cone = pd.read_excel(file_path, sheet_name="Cone")

df_cone_regija = df_cone[df_cone["Geo_enota"] == regija_name]

if df_cone_regija.empty:
    raise ValueError("Ni podatkov za izbrano regijo v listu Cone.")

# --- PRIPRAVA PODATKOV ---
# --- REGIJA ---
df_cone_regija = df_cone_regija.rename(columns={"Test": "TestIme"})
df_cone_regija["Regija_Tvegani"] = df_cone_regija.iloc[:, 8] * 100
df_cone_regija["Regija_Mejni"] = df_cone_regija.iloc[:, 7] * 100

# --- SLOVENIJA ---
df_cone_slo = df_cone[df_cone["Geo_enota"] == "Slovenija"]
df_cone_slo = df_cone_slo.rename(columns={"Test": "TestIme"})
df_cone_slo["Slo_Tvegani"] = df_cone_slo.iloc[:, 8] * 100
df_cone_slo["Slo_Mejni"] = df_cone_slo.iloc[:, 7] * 100

# --- ANALIZA (Existing code) ---
df_sql3["Skupaj"] = df_sql3[["Zdravo", "Mejno", "Tvegano"]].sum(axis=1)

df_sql3["Analiza_Tvegani"] = df_sql3.apply(
    lambda r: (r["Tvegano"] / r["Skupaj"] * 100) if r["Skupaj"] > 0 else 0,
    axis=1
)

df_sql3["Analiza_Mejni"] = df_sql3.apply(
    lambda r: (r["Mejno"] / r["Skupaj"] * 100) if r["Skupaj"] > 0 else 0,
    axis=1
)

# --- MERGE V ENO TABELO ---
df_plot = df_sql3.merge(
    df_cone_regija[["TestIme", "Regija_Tvegani", "Regija_Mejni"]],
    on="TestIme",
    how="left"
).merge(
    df_cone_slo[["TestIme", "Slo_Tvegani", "Slo_Mejni"]],
    on="TestIme",
    how="left"
)

# --- KEY FIX: SORT BEFORE ASSIGNING DATA FOR GRAPH ---
df_plot["Order"] = df_plot["TestIme"].map(order_map)
df_plot = df_plot.sort_values("Order")
df_plot = df_plot.drop(columns=["Order"])

# Now extract the data for the graph from the SORTED dataframe
testi = df_plot["TestIme"]
analiza_tvegani = df_plot["Analiza_Tvegani"]
analiza_mejni = df_plot["Analiza_Mejni"]
slo_tvegani = df_plot["Slo_Tvegani"]
slo_mejni = df_plot["Slo_Mejni"]
regija_tvegani = df_plot["Regija_Tvegani"] # Added for clarity
regija_mejni = df_plot["Regija_Mejni"]     # Added for clarity

# --- TABELA 2 (Will now match because df_plot is already sorted) ---
doc.add_page_break()
doc.add_heading('Delež mejnih in tveganih rezultatov po testih', level=2)
doc.add_paragraph("Tabela prikazuje delež udeležencev meritev, ki imajo testni dosežek označen kot tvegan (rdeča cona) ali mejen (rumena cona). Delež zelene cone ni prikazan, se pa lahko izračuna kot razlika 100% - Tvegano – Mejno. ")

table2 = doc.add_table(rows=2, cols=7)
if pd.notna(table_style) and str(table_style).strip() != "":
    try:
        table2.style = str(table_style)
    except Exception:
        print(f"Opozorilo: stil '{table_style}' ne obstaja, uporabljen bo privzeti.")
        table2.style = 'Table Grid'
else:
    table2.style = 'Table Grid'
table2.autofit = True
table2.allow_autofit = True

# --- PRVA VRSTICA (GLAVNE SKUPINE) ---
hdr1 = table2.rows[0].cells

hdr1[0].text = "Test / %"
hdr1[1].text = "ZmigajSdv"
hdr1[3].text = "Regija"
hdr1[5].text = "Slovenija"

# merge celice za skupine
hdr1[1].merge(hdr1[2])
hdr1[3].merge(hdr1[4])
hdr1[5].merge(hdr1[6])

# --- DRUGA VRSTICA (PODSTOLPCI) ---
hdr2 = table2.rows[1].cells

hdr2[0].text = ""

hdr2[1].text = "Tvegano"
hdr2[2].text = "Mejno"

hdr2[3].text = "Tvegano"
hdr2[4].text = "Mejno"

hdr2[5].text = "Tvegano"
hdr2[6].text = "Mejno"

# poravnava glave
for row in table2.rows[:2]:
    for cell in row.cells:
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

# --- PODATKI ---
for _, r in df_plot.iterrows():
    row_cells = table2.add_row().cells

    row_cells[0].text = str(r["TestIme"])

    row_cells[1].text = str(round(r["Analiza_Tvegani"], 1))
    row_cells[2].text = str(round(r["Analiza_Mejni"], 1))

    row_cells[3].text = str(round(r["Regija_Tvegani"], 1))
    row_cells[4].text = str(round(r["Regija_Mejni"], 1))

    row_cells[5].text = str(round(r["Slo_Tvegani"], 1))
    row_cells[6].text = str(round(r["Slo_Mejni"], 1))

    # poravnava
    row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
    for i in range(1, 7):
        row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
table2.columns[0].width = Cm(8.0)

for cell in table2.columns[0].cells:
    cell.width = Cm(8.0)
    for paragraph in cell.paragraphs:
        paragraph.paragraph_format.keep_together = True




# --- GRAF 2 ---
plt.figure(figsize=(12,6))

positions = list(range(len(testi)))
width = 0.18
offset = 0.25

# --- ANALIZA ---
x_analiza = [p - offset for p in positions]
plt.bar(x_analiza, analiza_tvegani, width=width, color='#df3127')
plt.bar(x_analiza, analiza_mejni, bottom=analiza_tvegani, width=width, color='#ffd54e')

# --- REGIJA (Updated to use the sorted variables) ---
x_regija = positions
plt.bar(x_regija, regija_tvegani, width=width, color='#df3127')
plt.bar(x_regija, regija_mejni, bottom=regija_tvegani, width=width, color='#ffd54e')

# --- SLO (Updated to use the sorted variables) ---
x_slo = [p + offset for p in positions]
plt.bar(x_slo, slo_tvegani, width=width, color='#df3127')
plt.bar(x_slo, slo_mejni, bottom=slo_tvegani, width=width, color='#ffd54e')

# --- NAVPIČNE ČRTE MED VSAKIM STOLPCEM ---

all_x = []
for i in range(len(testi)):
    all_x.extend([x_analiza[i], x_regija[i], x_slo[i]])

all_x_sorted = sorted(all_x)
# črta pred prvim stolpcem
first_line = all_x_sorted[0] - (all_x_sorted[1] - all_x_sorted[0]) / 2
plt.axvline(first_line, color='black', linewidth=0.5)

# črta za zadnjim stolpcem
last_line = all_x_sorted[-1] + (all_x_sorted[-1] - all_x_sorted[-2]) / 2
plt.axvline(last_line, color='black', linewidth=0.5)
for i in range(len(all_x_sorted) - 1):
    mid = (all_x_sorted[i] + all_x_sorted[i+1]) / 2
    plt.axvline(mid, color='black', linewidth=0.5)
def wrap_labels(labels, width=10):
    wrapped = []
    for label in labels:
        words = str(label).split()
        lines = []
        current = ""

        for w in words:
            if len(current) + len(w) + 1 <= width:
                current += (" " if current else "") + w
            else:
                lines.append(current)
                current = w
        if current:
            lines.append(current)

        wrapped.append("\n".join(lines))
    return wrapped
# --- X OZNAKE ---
wrapped_testi = wrap_labels(testi, width=10)

plt.xticks(positions, wrapped_testi, rotation=0, ha='center')

# pomakni imena testov dol
ax = plt.gca()
ax.tick_params(axis='x', pad=25)

# --- OZNAKE STOLPCEV (Z / R / S) ---
for i in range(len(testi)):
    plt.text(x_analiza[i], -3, "Z", ha='center', va='top', fontsize=9)
    plt.text(x_regija[i], -3, "R", ha='center', va='top', fontsize=9)
    plt.text(x_slo[i], -3, "S", ha='center', va='top', fontsize=9)

# osi
plt.ylabel("Odstotek")
plt.ylim(0, 100)

plt.tight_layout()

# shrani
graf2_path = os.path.join(os.getcwd(), "graf_cone.png")
plt.savefig(graf2_path)
plt.close()

# dodaj v DOCX
doc.add_paragraph("\n")
doc.add_picture(graf2_path, width=Inches(6))
doc.add_paragraph("Oznake skupin: Z=ZmigajSdv, R=Regija, S=Slovenija")
# --- SHRANJEVALNJE DOCX ---
docx_filename = os.path.join(os.getcwd(), ReportDocName + ".docx")
doc.save(docx_filename)
print(f"DOCX poročilo shranjeno v: {docx_filename}")
