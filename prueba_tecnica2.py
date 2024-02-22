# Dependencias:
# - pandas: Biblioteca para manipulación y análisis de datos en Python. Se utiliza para trabajar con DataFrames.
#   Instalación: pip install pandas

# - selenium: Herramienta para la automatización de navegadores web. Se utiliza para interactuar con la página web y realizar web scraping.
#   Instalación: pip install selenium

# Asegúrate de tener Google Chrome instalado antes de ejecutar el script.

import pandas as pd
import zipfile
import os
from io import BytesIO
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains

# Definir constantes
BASE_URL = "https://cde.ucr.cjis.gov/LATEST/webapp/#/pages/home"
DOWNLOADS_PAGE_URL = "https://cde.ucr.cjis.gov/LATEST/webapp/#/pages/downloads"
DOWNLOAD_BUTTON_ID = "nibrs-download-button"
TABLE_SELECT_ID = "dwnnibrs-download-select"
LOCATION_SELECT_ID = "dwnnibrsloc-select"
ZIP_FILENAME = "victims.zip"
XLSX_FILENAME = "Victims_Age_by_Offense_Category_2022.xlsx"
TARGET_CATEGORY = "Crimes Against Property"
CSV_FILENAME = "Crimes_Against_Property_2022.csv"

# Paths
PROJECT_PATH = os.path.dirname(os.path.abspath(__file__))
ZIP_LOCAL_PATH = os.path.join(PROJECT_PATH, ZIP_FILENAME)

def initialize_webdriver():
    options = Options()
    options.add_experimental_option("prefs", {"download.default_directory": os.getcwd()})
    return webdriver.Chrome(options=options)

def navigate_to_downloads_page(driver):
    driver.get(BASE_URL)
    download_go_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "home-dwnload-go-btn"))
    )
    download_go_button.click()
    WebDriverWait(driver, 10).until(
        EC.url_to_be(DOWNLOADS_PAGE_URL)
    )

def select_option_from_dropdown(driver, dropdown_id, option_text):
    dropdown_element = driver.find_element(By.ID, dropdown_id)
    ActionChains(driver).click(dropdown_element).perform()

    overlay = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, ".option-list"))
    )
    WebDriverWait(overlay, 10).until(
        EC.presence_of_all_elements_located((By.TAG_NAME, "nb-option"))
    )

    for option in overlay.find_elements(By.TAG_NAME, "nb-option"):
        if option.text == option_text:
            ActionChains(driver).click(option).perform()
            break

def download_file(driver, download_button_id, file_path):
    download_button = driver.find_element(By.ID, download_button_id)
    download_button.click()
    WebDriverWait(driver, 10).until(lambda d: os.path.exists(file_path))

def process_excel_and_generate_csv(zip_file_path, xlsx_filename, target_category):
    with zipfile.ZipFile(zip_file_path) as zip_file:
        if xlsx_filename not in zip_file.namelist():
            raise FileNotFoundError(f"El archivo {xlsx_filename} no está presente en el zip.")

        with pd.ExcelFile(BytesIO(zip_file.read(xlsx_filename))) as xls:
            sheet_name = xls.sheet_names[0]
            df = pd.read_excel(xls, sheet_name, header=[0, 1], skiprows=3)

            df_filtered = df[df[("Offense Category", "Unnamed: 0_level_1")] == target_category]

            csv_filename = "Crimes_Against_Property_2022.csv"
            df_filtered.to_csv(csv_filename, index=False)

            custom_header = ["Offense Category"] + ["Age: " + str(age[1]) for age in df.columns[2:]]

            df_csv = pd.read_csv(csv_filename, skiprows=[1])
            df_csv = df_csv.loc[:, ~df_csv.columns.str.contains('Total')]
            df_csv.columns = custom_header
            df_csv.to_csv(csv_filename, index=False)

            print(f"Se ha generado el archivo CSV: {csv_filename}")

try:
    # Inicializar el WebDriver
    driver = initialize_webdriver()

    # Navegar a la página de descargas
    navigate_to_downloads_page(driver)

    # Seleccionar opciones del formulario
    select_option_from_dropdown(driver, "dwnnibrs-download-select", "Victims")
    select_option_from_dropdown(driver, "dwnnibrsloc-select", "Florida")

    # Descargar el archivo
    download_file(driver, DOWNLOAD_BUTTON_ID, ZIP_LOCAL_PATH)

    # Procesar el archivo descargado y generar el CSV
    process_excel_and_generate_csv(ZIP_LOCAL_PATH, XLSX_FILENAME, TARGET_CATEGORY)

except FileNotFoundError as e:
    print(f"Error: {e}")
except Exception as e:
    print(f"Ocurrió un error inesperado: {e}")
finally:
    if 'driver' in locals() or 'driver' in globals():
        driver.quit()
