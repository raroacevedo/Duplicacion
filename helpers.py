from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from openpyxl import load_workbook, Workbook
from openpyxl.styles import NamedStyle

import warnings
warnings.simplefilter("ignore", category=UserWarning)  # Ocultar warnings

import pandas as pd
import openpyxl
import re

"""The brightspace_login method receives three arguments: the chromedriver, the user and the password of the Virtual Campus admin.
# If the login is succesful, it return True, otherwise, it returns False.
"""
path="courses.xlsx"
unidades_path="UnidadesOrganización.xlsx"

def brightspace_login(driver):
    try:
        print("\n")
        print("====================================================")
        print("        INICIANDO BOT DE DUPLICADO DE CURSOS        ")
        print("====================================================")
        # get the login page
        driver.get("https://virtual.upb.edu.co/d2l/login?noRedirect=1")

        # get the input elements
        username = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.ID, "userName"))
        )
        password = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.ID, "password"))
        )

        # send the credentials
        username.send_keys("adminbot")
        #password.send_keys(pwd)
        #password.send_keys("2/2*2Hijosmemsam")
        password.send_keys("*Upb/ÑV1rtu4l*")
        password.send_keys(Keys.RETURN)
        sleep(3)
        
        
        #2FA
        """driver.get("https://virtual.upb.edu.co/d2l/lp/auth/twofactorauthentication/TwoFactorCodeEntry.d2l")
        
        l2fa = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.ID, "z_d"))
        )
        l2fa.send_keys(dfa)
        l2fa.send_keys(Keys.RETURN)"""

        return True
    except:
        print(
            "Hubo un error al tratar de iniciar sesión. Revise sus credenciales e intente de nuevo."
        )

        return False


"""The create_chrome_driver receives no arguments. It creates the Chrome Driver with options to run on headless mode.
It returns the created driver.
"""


def create_chrome_driver():
    # service object
    #s = Service(r".\ChromeU\chromedriver-win64\chromedriver.exe")
    s = Service(r"..\Chrome\chromedriver.exe")

    # chrome options to run on headless mode and no logging on the terminal
    chrome_options = Options()
    chrome_options.add_argument("--disable-extensions")
    # chrome_options.add_argument("--headless")  # Si deseas correr en modo headless

    # Ocultar mensajes de DevTools y logs en la terminal
    chrome_options.add_argument("--log-level=3")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)

    # Desactivar el gestor de contraseñas de Chrome
    chrome_prefs = {
        "credentials_enable_service": False,  # Desactiva el guardado automático de contraseñas
        "profile.password_manager_enabled": False  # Desactiva el gestor de contraseñas
    }
    chrome_options.add_experimental_option("prefs", chrome_prefs)

    # Instanciar el webdriver
    driver = webdriver.Chrome(service=s, options=chrome_options)
    driver.implicitly_wait(40)

    return driver


""" The check_courses_file receives a single parameter, courses, which is a pandas dataframe that contains the courses we want to check in D2L Brightspace.
It checks that the file is in the correct form, returns True if it is, False otherwise.
"""

def check_courses_file(courses):
    # get the dataframe columns
    columns = list(courses.columns)

    # check if the number of columns correponds
    if len(columns) != 1:
        return False

    return "Enlace curso" in columns


"""The read_courses_url receives the file path containint the URLs. It checks if the file is ok, and returns the dataframe containint the information."""


def read_courses_url(course_file_path):
    try:
        # read the courses
        courses_url = pd.read_csv(course_file_path)

        # check if the file is ok
        courses_flag = check_courses_file(courses_url)
        if not courses_flag:
            return False

        return courses_url
    except:
        print(
            "Hubo un error leyendo el archivo de la URL de los cursos. Revise sus entradas e intente de nuevo."
        )
        return


"""The get_course_shortnames method receives two arguments: the chrome driver and the path to the file containing the course URLs. 
It writes a new file called shortnames.csv that contains the shortnames of the courses.
"""

def get_course_shortnames(driver, courses_file_path):
    course_info = []

    # read the courses URLs
    data = read_courses_url(courses_file_path)
    data_length = len(data)

    # loop through the courses
    for index, row in data.iterrows():
        print("Revisando el curso " + str(index + 1) + " de " + str(data_length))

        # get the current URL and ID
        current_course_url = row["Enlace curso"]
        id = current_course_url[-5:]

        # check if the ID is correct
        if id.isdigit():
            new_url = (
                "https://virtual.upb.edu.co/d2l/lp/manageCourses/course_offering_info_viewedit.d2l?ou="
                + id
            )
        else:
            id = current_course_url[-4:]
            new_url = (
                "https://virtual.upb.edu.co/d2l/lp/manageCourses/course_offering_info_viewedit.d2l?ou="
                + id
            )

        try:
            # go to the course
            driver.get(new_url)

            # get the course short name field and value
            short_name = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.ID, "z_l"))
            )
            short_name_value = short_name.get_attribute("value")

            # append the current course to the resulting dataframe
            course_info.append((short_name_value, short_name_value[-5:]))
        except Exception as e:
            print(e)

    # write the resulting file
    resulting_shortnames = pd.DataFrame(course_info, columns=["Nombre", "NRC"])
    resulting_shortnames.to_csv("shortnames.csv", index=False)

    return True


""" The check_courses_file receives a single parameter, courses, which is a pandas dataframe that contains
# the courses we want to duplicate in D2L Brightspace.
# It checks that the file is in the correct form, returns True if it is, False otherwise.
"""


def check_courses_file(courses):
    # get the dataframe columns
    columns = list(courses.columns)

    if len(columns) != 10:
        return False

    return (
        "Maestro" in columns
        and "Nombre" in columns
        and "Codigo" in columns
        and "Plantilla" in columns
        and "Semestre" in columns
        and "Oferta de curso - Id grupo Brightspace" in columns
        and "El maestro es QM" in columns
        and "Ruta de plantilla" in columns
        and "Se actualizo use lesson experience" in columns
        and "Se tuvo que actualizar content template path" in columns
    )


""" The duplicate_course method duplicates a course in D2L Brightspace with the info provided by the master file.
# It receives two parameters: drivers, which is the chromedriver, and the course object.
# Returns True if the course was duplicated, False otherwise.
"""

def select_course_template(driver, course):
    # STEP 1. COURSE TEMPLATE
    # find the select course template item, and wait for it to load
    templateSelect = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located(
            (
                By.CSS_SELECTOR,
                "#oldCTId_id",
            )
        )
    )
    templateSelectObject = Select(templateSelect)

    # select the given course template by value
    templateSelectObject.select_by_value(str(course["Plantilla"]))

    # wait for the next button and then click it
    nextBtn = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located(
            (
                By.CSS_SELECTOR,
                "body > form > table > tbody > tr > td > table > tbody > tr:nth-child(5) > td:nth-child(2) > table > tbody > tr > td:nth-child(2) > input",
            )
        )
    )
    nextBtn.click()
    sleep(2)  # wait for the next page to load


def course_offering_details(driver, course):
    # STEP 2. COURSE OFFERING DETAILS
    # wait for the course offering name input field and then fill it
    courseOfferingName = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located(
            (
                By.CSS_SELECTOR,
                "#courseOfferingName_id",
            )
        )
    )
    courseOfferingName.send_keys(course["Nombre"])

    # wait for the course offering code input field and then fill it
    courseOfferingCode = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located(
            (
                By.CSS_SELECTOR,
                "#courseOfferingCode_id",
            )
        )
    )
    courseOfferingCode.send_keys(course["Codigo"])

    # find the select course template item, and wait for it to load
    semesterSelect = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located(
            (
                By.CSS_SELECTOR,
                "#semId_id",
            )
        )
    )
    semesterSelectObject = Select(semesterSelect)

    # select the given course template by value
    semesterSelectObject.select_by_value(str(course["Semestre"]))

    # find the next button and wait for it to load
    courseNextBtn = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located(
            (
                By.CSS_SELECTOR,
                "body > form > table > tbody > tr > td > table > tbody > tr:nth-child(5) > td:nth-child(2) > table > tbody > tr > td:nth-child(3) > input",
            )
        )
    )
    courseNextBtn.click()
    sleep(4)

# The confirm_course_creation method receives the chromedriver and confirms the course creation.
# It waits for the confirm button to load, and then clicks it.
def confirm_course_creation(driver):
    # STEP 3. CONFIRM COURSE CREATION
    # wait for the confirm button to load, and then click it
    confirmCourseBtn = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located(
            (
                By.CSS_SELECTOR,
                "body > form > table > tbody > tr > td > table > tbody > tr:nth-child(5) > td:nth-child(2) > table > tbody > tr > td:nth-child(3) > input",
            )
        )
    )
    confirmCourseBtn.click()
    sleep(4)

# The import_course_content method receives the chromedriver, the course object and the index of the course in the dataframe.
# It imports the course content from the master course, and updates the excel file with the Bright
def import_course_content(driver, course, index):
    # save the current page for window handling
    mainPage = None
    while not mainPage:
        mainPage = driver.current_window_handle

    # STEP 4. COPY COURSE COMPONENTS

    # click the link to copy course components
    copyCourseComponents = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located(
            (
                By.CSS_SELECTOR,
                "body > form > table > tbody > tr > td > table > tbody > tr:nth-child(4) > td:nth-child(3) > table:nth-child(8) > tbody > tr > td > li:nth-child(2) > a",
            )
        )
    )
    sleep(2)

    copyCourseComponents.click()
    sleep(2)

    # click the button to search for the master course
    searchOfferBtn = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located(
            (
                By.CSS_SELECTOR,
                "#z_j",  #cambio el 20/nov/2024
            )
        )
    )
    searchOfferBtn.click()
    sleep(2)

    #changing the handles to access search page
    #for handle in driver.window_handles:
    #    print (" popup - "+str(handle))
    searchPage = None
    while not searchPage:
        for handle in driver.window_handles:
            if handle != mainPage:
                searchPage = handle
                break

    # change the control to search page
    driver.switch_to.window(searchPage)
    sleep(2)

    #find the frame containing the search bar for the course offering
    #frame = driver.find_element(By.CSS_SELECTOR, "#PopupWindow > frame:nth-child(3)")   --- sin UserWay
    #frame = driver.find_element(By.XPATH, "/html/frameset/frame[2]")
    #frame = driver.find_element(By.CSS_SELECTOR, "#PopupWindow > frame:nth-child(10)")  -- cambio 2023/11/15
    frame = driver.find_element(By.CSS_SELECTOR, "frame[title='Cuerpo']")

    driver.switch_to.frame(frame)
    sleep(2)

    #searchInput = driver.find_element(By.XPATH, '//input[@id="z_b"]')
    searchInput = driver.find_element(By.CSS_SELECTOR, "#z_b")
    
    searchInput.send_keys(str(course["Maestro"]))
    searchInput.send_keys(Keys.RETURN)
  
    driver.find_element(
        By.CSS_SELECTOR,
        "#yui-rec0 > td.d_dg_col_d_selection.yui-dt0-col-d_selection.yui-dt-col-d_selection.yui-dt-first > div > span > input",
    ).click()
    sleep(2)  # wait for the next page to load

    # find the frame containing the confirmation button
    driver.switch_to.window(searchPage)
    #newFrame = driver.find_element(By.CSS_SELECTOR, "#PopupWindow > frame:nth-child(4)")  --- sin UserWay
    #newFrame = driver.find_element(By.CSS_SELECTOR, "#PopupWindow > frame:nth-child(11)")  cambio 2023/11715
    newFrame = driver.find_element(By.CSS_SELECTOR, "frame[title='Pie de página']")         
    driver.switch_to.frame(newFrame)
    driver.find_element(By.CSS_SELECTOR, "#z_a > div > button:nth-child(1)").click()
    driver.switch_to.window(mainPage)
    driver.find_element(By.CSS_SELECTOR, "#z_b").click()
    sleep(25)
    
    # Obtener la URL actual y extraer el ID del curso
    current_url = driver.current_url
    match = re.search(r'copy/(\d+)/History', current_url)
    if match:
        brightspace_id = match.group(1)
        print(f"\nID de Brightspace obtenido: {brightspace_id}")
    else:
        print("\nNo se encontró el ID de Brightspace en la URL")
        return

    # Escribir el ID en Excel sin perder formatos
    update_excel(path, index, brightspace_id)

# The update_excel method receives the file path, the row index and the brightspace_id.
# It updates the column 'Oferta de curso - Id grupo Brightspace' without losing formats and
def update_excel(file_path, row_index, brightspace_id):
    
    """actualiza la columna 'Oferta de curso - Id grupo Brightspace' sin perder formatos y sin afectar otras lineas"""

    # Intenta cargar el archivo Excel
    try:
        wb = load_workbook(file_path)
    except Exception as e:
        print(f"\nError al abrir el archivo: {e}")
        return

    ws = wb.active  # Seleccionar la hoja activa

    # Crear un estilo predeterminado solo si el libro no lo tiene
    if "default" not in wb.named_styles:
        default_style = NamedStyle(name="default")
        wb.add_named_style(default_style)
        
    brightspace_col = 6  # Columna F es la número 6 (1-based)
    excel_row = row_index + 2  # +2 porque pandas indexa desde 0 y Excel tiene encabezado

    try:
        ws.cell(row=excel_row, column=brightspace_col, value=int(brightspace_id))
        print(f"\nID {brightspace_id} agregado a la fila {excel_row} correctamente")
    except Exception as e:
        print(f"\nError actualizando la celda: {e}")


    # Guardar los cambios en un archivo temporal y luego reemplazar el original
    temp_path = file_path.replace(".xlsx", "_temp.xlsx")
    wb.save(temp_path)
    wb.close()

    # Reemplazar el archivo original con el nuevo
    import os
    os.replace(temp_path, file_path)
   
    
# The new_experience method receives the chromedriver and a row of the courses dataframe.
# It checks if the course is a QM course, and if it is, it updates the
def new_experience(driver, row):
    
    # Cargar el archivo UnidadesOrganización.xlsx
    unidades_df = pd.read_excel("UnidadesOrganización.xlsx")

    # Buscar el maestro en UnidadesOrganización
    maestro_value = row["Maestro"]
    unidad = unidades_df[unidades_df["Code"] == maestro_value]

    if not unidad.empty:
        name_value = str(unidad.iloc[0]["Name"]).strip()  # Convertir a texto limpio
        qm_status = "SI" if "QM" in name_value else "NO"

        print(f"\nMaestro: {maestro_value} | Nombre: {name_value} | QM: {qm_status}")

        # Cargar el archivo Excel sin perder formatos
        book = load_workbook(path)
        sheet = book.active  

        # Obtener la fila real en Excel (ajuste +2 por header)
        row_index = row.name + 2
        column_index_qm = 7  # Columna "El maestro es QM"

        # Escribir "SI" o "NO" en la celda correspondiente
        sheet.cell(row=row_index, column=column_index_qm, value=qm_status)

        # Determinar la ruta de plantilla basada en "El maestro es QM"
        if qm_status == "SI":
            ruta_plantilla = "/shared/HTML-Template-QM/"
        elif qm_status == "NO":
            ruta_plantilla = "/shared/HTML-Template-Library/"
        else:
            ruta_plantilla = ""  # En caso de que no tenga un valor válido
        # Escribir en la celda de "Ruta de plantilla" asegurando que sea texto
        sheet.cell(row=row_index, column=8).value = str(ruta_plantilla)

        # Guardar cambios en el Excel
        book.save(path)
        book.close()

        print(f"\nRuta de plantilla actualizada en el excel con: {ruta_plantilla}")

        # Recargar el DataFrame después de la modificación con openpyxl
        df_actualizado = pd.read_excel(path, dtype={"Oferta de curso - Id grupo Brightspace": str})
        row_actualizado = df_actualizado.iloc[row.name]

        # Ahora el código de oferta y la ruta de plantilla deben estar correctos
        codigo_oferta = str(row_actualizado["Oferta de curso - Id grupo Brightspace"]).strip()
        ruta_plantilla = str(row_actualizado["Ruta de plantilla"]).strip()

        # Validación para evitar errores
        if codigo_oferta.lower() == "nan" or codigo_oferta == "":
            print(f"\nAdvertencia: Código de oferta sigue siendo NaN. Se omite esta actualización.")
            return  # Salir de la función

        if ruta_plantilla.lower() == "nan" or ruta_plantilla == "":
            print(f"\nAdvertencia: Ruta de plantilla sigue siendo NaN. Se omite esta actualización.")
            return  # Salir de la función


        if qm_status == "SI":
            try:
                # Acceder y actualizar en la página web con Selenium
                driver.get("https://virtual.upb.edu.co/d2l/lp/configVariableBrowser?configPath=d2l.Tools.Content.UseLessonsExperience")
                sleep(3)

                # Botón agregar valor
                wait = WebDriverWait(driver, 10)
                boton = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Agregar valor']")))
                boton.click()
                sleep(3)

                # Cambiar a iframe correcto
                iframes = driver.find_elements(By.TAG_NAME, "iframe")
                for iframe in iframes:
                    if "d2l-dialog-frame" in iframe.get_attribute("class"):
                        driver.switch_to.frame(iframe)
                        break  

                # Ingresar código en el campo de la web
                campo = driver.find_element(By.NAME, "orgUnitIdEdit")  
                campo.clear()
                campo.send_keys(codigo_oferta)
                sleep(3)

                # Seleccionar la opción "activada"
                select_element = driver.find_element(By.NAME, "input$param")
                select = Select(select_element)
                select.select_by_value("true")
                sleep(3)

                # Clic en guardar
                guardar_boton = driver.find_element(By.XPATH, "//button[contains(@class, 'd2l-button') and text()='Guardar']")
                guardar_boton.click()
                sleep(3)

                driver.switch_to.default_content()

                # Escribir en Excel que se actualizó (Forma correcta)
                book = load_workbook(path)
                sheet = book.active
                sheet.cell(row=row_index, column=9, value="SI")  # Columna "Se actualizo use lesson experience"
                book.save(path)
                book.close()

                print("\nActualización de Use Lessons Experience completada, se actualiza el excel con SI.")

                # Repetir proceso para "Content Template Path"
                driver.get("https://virtual.upb.edu.co/d2l/lp/configVariableBrowser?configPath=Content.DefaultTemplateDirectory")
                sleep(3)

                boton = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Agregar valor']")))
                boton.click()
                sleep(3)

                iframes = driver.find_elements(By.TAG_NAME, "iframe")
                for iframe in iframes:
                    if "d2l-dialog-frame" in iframe.get_attribute("class"):
                        driver.switch_to.frame(iframe)
                        break  

                campo = driver.find_element(By.NAME, "orgUnitIdEdit")  
                campo.clear()
                campo.send_keys(codigo_oferta)
                sleep(2)

                campo_input = driver.find_element(By.NAME, "input$param")  
                campo_input.clear()
                campo_input.send_keys(ruta_plantilla)
                sleep(2)

                guardar_boton = driver.find_element(By.XPATH, "//button[contains(@class, 'd2l-button') and text()='Guardar']")
                guardar_boton.click()
                sleep(3)

                driver.switch_to.default_content()

                # Escribir en Excel que se actualizó "Content Template Path"
                book = load_workbook(path)
                sheet = book.active
                sheet.cell(row=row_index, column=10, value="SI")  # Columna "Se tuvo que actualizar content template path"
                book.save(path)
                book.close()

                print("\nActualización de Content Template Path completada, se actualiza el excel con SI.")

            except Exception as e:
                print(f"\nError al procesar el curso {row['Nombre']}: {e}")

        else:
            # Si es "NO", actualizar las columnas en Excel con "NO"
            book = load_workbook(path)
            sheet = book.active
            sheet.cell(row=row_index, column=9, value="NO")  # Columna "Se actualizo use lesson experience"
            sheet.cell(row=row_index, column=10, value="NO")  # Columna "Se tuvo que actualizar content template path"
            book.save(path)
            book.close()

            print(f"\nCurso {row['Nombre']}: No requiere actualización de experiencia. Se actualiza el Excel con 'NO'.")
    else:
        print(f"\n Maestro {maestro_value} no encontrado en el archivo UnidadesOrganización.xlsx. Se omite la actualización de experiencia.")
            

def duplicate_course(driver, course, index):
    # access the course duplication page
    driver.get(
        "https://virtual.upb.edu.co/d2l/tools/courseCreate/courseCreateType.asp?ou=6606"
    )

    # STEP 1
    select_course_template(driver, course)

    # STEP 2
    course_offering_details(driver, course)

    # STEP 3
    confirm_course_creation(driver)

    # STEP 4
    import_course_content(driver, course, index)
    
    # STEP 5
    new_experience(driver, course)

    return
