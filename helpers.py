from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import os
import logging
from openpyxl import load_workbook, Workbook
from openpyxl.styles import NamedStyle

import warnings
warnings.simplefilter("ignore", category=UserWarning)  # Ocultar warnings

import pandas as pd
import openpyxl
import re


logger = logging.getLogger(__name__)

# Funciones de normalización para extraer NRC y período desde el código del curso
def normalize_banner_nrc(value):
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none"}:
        return ""
    if text.endswith(".0"):
        text = text[:-2]
    return text

# El período en Banner a menudo viene con formato "202440.0" o similar, 
# así que esta función limpia el texto y extrae solo los dígitos relevantes
def normalize_banner_periodo(value):
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none"}:
        return ""
    try:
        return str(int(float(text)))
    except (TypeError, ValueError):
        digits = "".join(ch for ch in text if ch.isdigit())
        return digits if digits else text

"""The brightspace_login method receives three arguments: the chromedriver, the user and the password of the Virtual Campus admin.
# If the login is succesful, it return True, otherwise, it returns False.
"""
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
        
        
        #2FA ESTE USUARIO NO TIENE 2FA, PERO DEJÓ EL CÓDIGOS POR SI SE NECESITA IMPLEMENTAR EN EL FUTURO
        """driver.get("https://virtual.upb.edu.co/d2l/lp/auth/twofactorauthentication/TwoFactorCodeEntry.d2l")
        
        l2fa = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.ID, "z_d"))
        )
        l2fa.send_keys(dfa)
        l2fa.send_keys(Keys.RETURN)"""

        return True
    except Exception:
        logger.exception("Error al tratar de iniciar sesión en Brightspace.")
        print(
            "Hubo un error al tratar de iniciar sesión. Revise sus credenciales e intente de nuevo."
        )

        return False


"""The create_chrome_driver receives no arguments. It creates the Chrome Driver with options to run on headless mode.
It returns the created driver.
"""

def create_chrome_driver():
    """Configura y retorna el WebDriver con opciones seguras sin el prompt de red local"""
    service = Service(r"..\Chrome\chromedriver.exe")

    # Verifica si el chromedriver existe en la ruta especificada
    if not os.path.exists(service.path):    
        raise FileNotFoundError(f"El chromedriver no se encuentra en la ruta: {service.path}")
    
    # Configuración de opciones del navegador
    options = Options()
    options.add_argument("--disable-extensions")
    options.add_argument("--incognito")              
    options.add_argument("--disable-notifications")  
    options.add_argument("--allow-insecure-localhost") 
    options.add_argument("--log-level=3")
    
    # --- NUEVOS PARÁMETROS PARA OMITIR EL MENSAJE DE RED LOCAL ---
    # 1. Deshabilita la característica que hace saltar el aviso de seguridad
    options.add_argument("--disable-features=LocalNetworkAccessChecks")
    
    # 2. Inyecta preferencias directas al perfil para autoconceder el permiso si llega a ser requerido
    prefs = {
        "profile.managed_default_content_settings.local_network_access": 1
    }
    options.add_experimental_option("prefs", prefs)
    # -------------------------------------------------------------

    driver = webdriver.Chrome(service=service, options=options)
    driver.implicitly_wait(10)
    
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
    except Exception:
        logger.exception("Error leyendo el archivo de URL de cursos: %s", course_file_path)
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
            logger.exception("Error obteniendo shortname para URL %s", current_course_url)
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

# STEP 1. COURSE TEMPLATE
def select_course_template(driver, course):
    # find the select course template item, and wait for it to load
    wait = WebDriverWait(driver, 15)

    # 1) Host 1 (shadow root 0)
    host1 = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#LitId")))
    shadow0 = host1.shadow_root

    # 2) Host 2 (shadow root 1)
    host2 = wait.until(lambda d: shadow0.find_element(By.CSS_SELECTOR, "d2l-create-course-view"))
    shadow1 = host2.shadow_root

    # 3) Host 3 (el botón; OJO: d2l-button-subtle ES el host clickable)
    btn_host = wait.until(lambda d: shadow1.find_element(By.CSS_SELECTOR, "d2l-button-subtle[type='button']"))
    btn_shadow = btn_host.shadow_root

    # Busca un <button> interno o un slot clickeable (dependiendo del componente)
    inner_btn = wait.until(lambda d: btn_shadow.find_element(By.CSS_SELECTOR, "button, [role='button']"))

    try:
        inner_btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", inner_btn)
    
    #
    # Obtener el valor de la plantilla a seleccionar desde el DataFrame y asignarlo a una variable, asegurando que sea texto limpio
    #
    valorplantilla = str(course["Plantilla"])

    # Host 1
    host1 = wait.until(lambda d: d.find_element(By.CSS_SELECTOR, "#LitId"))
    shadow0 = host1.shadow_root

    # Host 2
    host2 = wait.until(lambda d: shadow0.find_element(By.CSS_SELECTOR, "d2l-create-course-view"))
    shadow1 = host2.shadow_root

    # Host 3
    host3 = wait.until(lambda d: shadow1.find_element(By.CSS_SELECTOR, "d2l-create-course-change-template-view"))
    shadow2 = host3.shadow_root

    # Host 4
    host4 = wait.until(lambda d: shadow2.find_element(By.CSS_SELECTOR, "#searchTemplate"))
    shadow3 = host4.shadow_root

    # Host 5 (d2l-input-text)
    input_host = wait.until(
        lambda d: shadow3.find_element(By.CSS_SELECTOR, "d2l-input-text[placeholder='Buscar plantillas de cursos']")
    )
    input_shadow = input_host.shadow_root

    # Buscar el input real dentro del shadow del componente
    real_input = wait.until(lambda d: input_shadow.find_element(By.CSS_SELECTOR, "input, textarea"))

    # Limpiar y escribir
    real_input.click()
    real_input.send_keys(Keys.CONTROL, "a")
    real_input.send_keys(Keys.BACKSPACE)
    real_input.send_keys(valorplantilla)
    real_input.send_keys(Keys.ENTER)

    sleep(2)  # Esperar a que se carguen los resultados de la búsqueda 
  
    #
    # Buscar el botón de selección (un elemento con role="button") buscar plantilla
    #
    # 1) Host 1
    """ Con el enter se elimina la necesidad de hacer click en el botón de búsqueda, pero dejo este código comentado por si se necesita hacer click manualmente (dependiendo de la velocidad de carga, a veces el enter no funciona porque el botón no está listo)
   
    # 2) Host 2
    host1 = wait.until(lambda d: d.find_element(By.CSS_SELECTOR, "#LitId"))
    shadow0 = host1.shadow_root
    

    host2 = wait.until(lambda d: shadow0.find_element(By.CSS_SELECTOR, "d2l-create-course-view"))
    shadow1 = host2.shadow_root

    # 3) Host 3
    host3 = wait.until(lambda d: shadow1.find_element(By.CSS_SELECTOR, "d2l-create-course-change-template-view"))
    shadow2 = host3.shadow_root

    # 4) Host 4
    host4 = wait.until(lambda d: shadow2.find_element(By.CSS_SELECTOR, "#searchTemplate"))
    shadow3 = host4.shadow_root

    # 5) Botón icono (HOST)
    btn_host = wait.until(lambda d: shadow3.find_element(By.CSS_SELECTOR, "d2l-button-icon[type='button']"))

    # Click normal o fallback JS
    try:
        btn_host.click()
    except Exception:
        driver.execute_script("arguments[0].click();", btn_host)
     """
    
    #
    #seleccionar el primer resultado (asumiendo que es el correcto)
    #
    # 1) Host 1
    #host1 = wait.until(lambda d: d.find_element(By.CSS_SELECTOR, "#LitId"))
    #shadow0 = host1.shadow_root

    # 2) Host 2
    host2 = wait.until(lambda d: shadow0.find_element(By.CSS_SELECTOR, "d2l-create-course-view"))
    shadow1 = host2.shadow_root

    # 3) Host 3
    host3 = wait.until(lambda d: shadow1.find_element(By.CSS_SELECTOR, "d2l-create-course-change-template-view"))
    shadow2 = host3.shadow_root

    # 4) d2l-selection-input con key dinámico (ARMAMOS SELECTOR)
    sel_css = f"d2l-selection-input[key='{valorplantilla}']"
    sel_host = wait.until(lambda d: shadow2.find_element(By.CSS_SELECTOR, sel_css))

    # 5) Click al checkbox id d2l-uid-*)
    try:
        sel_host.click()
    except Exception:
        driver.execute_script("arguments[0].click();", sel_host)
    
    #
    # Dar cklick en el botón de guardar (que también está dentro de shadow DOMs)
    # Host 1
    #host1 = wait.until(lambda d: d.find_element(By.CSS_SELECTOR, "#LitId"))
    #shadow0 = host1.shadow_root

    # Host 2
    host2 = wait.until(lambda d: shadow0.find_element(By.CSS_SELECTOR, "d2l-create-course-view"))
    shadow1 = host2.shadow_root

    # Host 3 (diálogo)
    dlg_host = wait.until(lambda d: shadow1.find_element(
        By.CSS_SELECTOR,
        "d2l-create-course-dialog[dialog-title='Cambiar plantilla del curso']"
    ))
    dlg_shadow = dlg_host.shadow_root

    # Botón (tu ruta actual)
    save_btn = wait.until(lambda d: dlg_shadow.find_element(
        By.CSS_SELECTOR,
        "d2l-dialog-fullscreen:nth-child(1) > d2l-button:nth-child(2)"
    ))

    try:
        save_btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", save_btn)

    sleep(2)  # wait for the next page to load


# STEP 2. COURSE SEMESTRE
def select_course_semestre(driver, course):
    #buscar el select de semestre y seleccionar el semestre correspondiente al curso
    wait = WebDriverWait(driver, 15)

    host1 = wait.until(lambda d: d.find_element(By.CSS_SELECTOR, "#LitId"))
    shadow0 = host1.shadow_root

    host2 = wait.until(lambda d: shadow0.find_element(By.CSS_SELECTOR, "d2l-create-course-view"))
    shadow1 = host2.shadow_root

    # Menos rígido: solo por el mensaje de validación
    host3 = wait.until(lambda d: shadow1.find_element(
        By.CSS_SELECTOR,
        "d2l-button-subtle-form-element[validation-error-dialog-message='Se requiere Semestre.']"
    ))
    shadow2 = host3.shadow_root

    btn_host = wait.until(lambda d: shadow2.find_element(By.CSS_SELECTOR, "d2l-button-subtle[type='button']"))

    try:
        btn_host.click()
    except Exception:
        driver.execute_script("arguments[0].click();", btn_host)

    #
    # Obtener el valor del semestre a seleccionar desde el DataFrame y asignarlo a una variable, asegurando que sea texto limpio
    #
    valor_semestre = str(course["Semestre"])

    # Host 1
    #host1 = wait.until(lambda d: d.find_element(By.CSS_SELECTOR, "#LitId"))
    #shadow0 = host1.shadow_root

    # Host 2
    #host2 = wait.until(lambda d: shadow0.find_element(By.CSS_SELECTOR, "d2l-create-course-view"))
    shadow1 = host2.shadow_root

    # Host 3
    sem_selector = wait.until(lambda d: shadow1.find_element(
        By.CSS_SELECTOR, "d2l-semester-selector[semester-ancestor-id='6606']"
    ))
    shadow2 = sem_selector.shadow_root

    # Host 4
    container = wait.until(lambda d: shadow2.find_element(
        By.CSS_SELECTOR, ".align-search-semester-for-inline-container"
    ))
    shadow3 = container.shadow_root

    # Host 5 (d2l-input-text)
    input_host = wait.until(lambda d: shadow3.find_element(
        By.CSS_SELECTOR, "d2l-input-text[placeholder='Buscar por {semestre}']"
    ))
    input_shadow = input_host.shadow_root

    # Input real dentro del webcomponent
    real_input = wait.until(lambda d: input_shadow.find_element(By.CSS_SELECTOR, "input, textarea"))

    # Limpiar y escribir
    real_input.click()
    real_input.send_keys(Keys.CONTROL, "a")
    real_input.send_keys(Keys.BACKSPACE)
    real_input.send_keys(valor_semestre)
    real_input.send_keys(Keys.ENTER) ##se hace enter para evitar el click en el botón de búsqueda

    sleep(2)  # Esperar a que se carguen los resultados de la búsqueda

    #
    #seleccionar el primer resultado (asumiendo que es el correcto)
    #
    # 1) Host 1
    #host1 = wait.until(lambda d: d.find_element(By.CSS_SELECTOR, "#LitId"))
    #shadow0 = host1.shadow_root

    # Host 1
    #host1 = wait.until(lambda d: d.find_element(By.CSS_SELECTOR, "#LitId"))
    #shadow0 = host1.shadow_root

    # Host 2
    #host2 = wait.until(lambda d: shadow0.find_element(By.CSS_SELECTOR, "d2l-create-course-view"))
    shadow1 = host2.shadow_root

    # Host 3
    sem_host = wait.until(lambda d: shadow1.find_element(
        By.CSS_SELECTOR, "d2l-semester-selector[semester-ancestor-id='6606']"
    ))
    shadow2 = sem_host.shadow_root

    # Espera a que haya al menos un resultado en la lista
    first_item = wait.until(lambda d: shadow2.find_elements(By.CSS_SELECTOR, "d2l-selection-input")[0])

    # Click directo al host (rápido, como tu Optimización 2)
    try:
        first_item.click()
    except Exception:
        driver.execute_script("arguments[0].click();", first_item)

    #
    # Dar cklick en el botón de guardar (que también está dentro de shadow DOMs)
    # Host 1
    #host1 = wait.until(lambda d: d.find_element(By.CSS_SELECTOR, "#LitId"))
    #shadow0 = host1.shadow_root

    # Host 2
    #host2 = wait.until(lambda d: shadow0.find_element(By.CSS_SELECTOR, "d2l-create-course-view"))
    shadow1 = host2.shadow_root

    # Host 3 (selector de semestre)
    sem_host = wait.until(lambda d: shadow1.find_element(
        By.CSS_SELECTOR, "d2l-semester-selector[semester-ancestor-id='6606']"
    ))
    shadow2 = sem_host.shadow_root

    # Host 4 (botón done) -> CLICK DIRECTO AL HOST
    done_btn = wait.until(lambda d: shadow2.find_element(
        By.CSS_SELECTOR, "d2l-button[type='button'][data-dialog-action='done']"
    ))

    try:
        done_btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", done_btn)

    sleep(2) # wait for the next page to load

#
# The course_offering_details method receives the chromedriver and a row of the courses dataframe.
#
def course_offering_details(driver, course, banner_lookup):
    # STEP 3. COURSE OFFERING DETAILS

    # Esperar a que cargue el formulario de detalles del curso (que también tiene shadow DOMs) y llenar los campos de nombre y código del curso
    wait = WebDriverWait(driver, 15)
    valor_nombre = str(course["Nombre"]) #nombre del curso a ingresar, asegurando que sea texto limpio

    # 1) Host raíz
    host1 = wait.until(lambda d: d.find_element(By.CSS_SELECTOR, "#LitId"))
    shadow0 = host1.shadow_root

    # 2) Vista crear curso
    host2 = wait.until(lambda d: shadow0.find_element(By.CSS_SELECTOR, "d2l-create-course-view"))
    shadow1 = host2.shadow_root

    # 3) Componente del nombre (host con id estable)
    name_host = wait.until(lambda d: shadow1.find_element(By.CSS_SELECTOR, "#course-name-input"))
    name_shadow = name_host.shadow_root

    # 4) Input real interno (estable)
    real_input = wait.until(lambda d: name_shadow.find_element(By.CSS_SELECTOR, "input, textarea"))

    # 5) Limpiar + escribir
    real_input.click()
    real_input.send_keys(Keys.CONTROL, "a")
    real_input.send_keys(Keys.BACKSPACE)
    real_input.send_keys(valor_nombre)
    
    #
    # wait for the course offering code input field and then fill it
    #
    valor_codigo = str(course["Codigo"]) #código del curso a ingresar, asegurando que sea texto limpio

    # 1) Host raíz
    #host1 = wait.until(lambda d: d.find_element(By.CSS_SELECTOR, "#LitId"))
    #shadow0 = host1.shadow_root

    # 2) Vista crear curso
    #host2 = wait.until(lambda d: shadow0.find_element(By.CSS_SELECTOR, "d2l-create-course-view"))
    shadow1 = host2.shadow_root

    # 3) Componente del código
    code_host = wait.until(lambda d: shadow1.find_element(By.CSS_SELECTOR, "#course-code-input"))
    code_shadow = code_host.shadow_root

    # 4) Input real interno
    real_input = wait.until(lambda d: code_shadow.find_element(By.CSS_SELECTOR, "input, textarea"))

    # 5) Limpiar + escribir
    real_input.click()
    real_input.send_keys(Keys.CONTROL, "a")
    real_input.send_keys(Keys.BACKSPACE)
    real_input.send_keys(valor_codigo)

    sleep(2)  # wait for the next page to load

    #
    # Agregar fecha de inicio y fin del curso datos se extraen desde Banner
    #

    #Desde este codigo se busca el NRC es el dato despues del ultimo "-" 
    #y se extrae el periodo desde el mismo codigo esta despeus del segundo guion, 
    #asegurar que es los primero 6 digitos, por ejemplo: CODIGO-202440-12345 -> periodo: 202440, NRC: 12345
    codigo_parts = valor_codigo.split("-")
    if len(codigo_parts) < 2:
        raise ValueError(
            f"El código del curso '{valor_codigo}' no tiene formato esperado para extraer período y NRC."
        )

    nrc = normalize_banner_nrc(codigo_parts[-1])
    periodo = normalize_banner_periodo(codigo_parts[-2][:6])
    if not nrc or not periodo:
        raise ValueError(
            f"No fue posible extraer NRC/período desde el código del curso '{valor_codigo}'."
        )

    # Buscar la fecha de inicio y fin en el lookup de Banner precargado
    banner_key = (nrc, periodo)

    banner_dates = banner_lookup.get(banner_key)
    if not banner_dates:
        raise ValueError(
            "No se encontró información en Banner para "
            f"NRC '{nrc}' y período '{periodo}' (clave buscada: {banner_key})."
        )

    fecha_inicio, fecha_fin = banner_dates

    #agregar la fecha de inicio 
    valor_fecha_inicio = str(fecha_inicio)
    valor_fecha_fin = str(fecha_fin)

    # 1) Host raíz
    host1 = wait.until(lambda d: d.find_element(By.CSS_SELECTOR, "#LitId"))
    shadow0 = host1.shadow_root

    # 2) Vista principal
    host2 = wait.until(lambda d: shadow0.find_element(By.CSS_SELECTOR, "d2l-create-course-view"))
    shadow1 = host2.shadow_root

    # 3) Componente de rango de fechas
    date_range_host = wait.until(lambda d: shadow1.find_element(By.CSS_SELECTOR, "#date-range-input"))

    # 4) Buscar todos los inputs profundos dentro del shadow DOM
    inputs = find_inputs_deep(driver, date_range_host)

    # 5) Filtrar inputs visibles/habilitados
    valid_inputs = []
    for inp in inputs:
        try:
            if inp.is_displayed() and inp.is_enabled():
                valid_inputs.append(inp)
        except Exception:
            continue

    if not valid_inputs:
        raise Exception(
            "No se encontraron inputs visibles y habilitados dentro de #date-range-input. "
            "El componente puede no haber terminado de renderizar o su estructura cambió."
        )

    # 6) Tomar el primer input como fecha inicio
    fecha_inicio_input = valid_inputs[0]

    # 6) Tomar el segundo input como fecha final
    fecha_fin_input = valid_inputs[1]

    # 7) Limpiar y escribir la fecha de inicio (con manejo de excepciones para inputs difíciles)
    try:
        fecha_inicio_input.click()
        fecha_inicio_input.send_keys(Keys.CONTROL, "a")
        fecha_inicio_input.send_keys(Keys.BACKSPACE)
        fecha_inicio_input.send_keys(valor_fecha_inicio)
    except Exception:
        driver.execute_script("""
            const el = arguments[0];
            const value = arguments[1];
            el.focus();
            el.value = value;
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
        """, fecha_inicio_input, valor_fecha_inicio)
    
    # 8) Limpiar y escribir la fecha de fin (con manejo de excepciones para inputs difíciles)
    try:
        fecha_fin_input.click()
        fecha_fin_input.send_keys(Keys.CONTROL, "a")
        fecha_fin_input.send_keys(Keys.BACKSPACE)
        fecha_fin_input.send_keys(valor_fecha_fin)
    except Exception:
        driver.execute_script("""
            const el = arguments[0];
            const value = arguments[1];
            el.focus();
            el.value = value;
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
        """, fecha_fin_input, valor_fecha_fin)

    sleep(2)

#buscar recursivamente dentro de un shadow DOM dado un host raíz, y devolver todos los elementos input 
# o textarea encontrados (incluso si están anidados dentro de múltiples niveles de shadow DOM)
def find_inputs_deep(driver, root_host):
    script = """
    function collectInputsDeep(root) {
        let results = [];

        function walk(nodeRoot) {
            if (!nodeRoot) return;

            const directInputs = nodeRoot.querySelectorAll('input, textarea');
            directInputs.forEach(el => results.push(el));

            const all = nodeRoot.querySelectorAll('*');
            for (const el of all) {
                if (el.shadowRoot) {
                    walk(el.shadowRoot);
                }
            }
        }

        walk(root_host.shadowRoot);
        return results;
    }

    const root_host = arguments[0];
    return collectInputsDeep(root_host);
    """
    return driver.execute_script(script, root_host)

# The confirm_course_creation method receives the chromedriver and confirms the course creation.
# It waits for the confirm button to load, and then clicks it.
def confirm_course_creation(driver):
    # STEP 4. CONFIRM COURSE CREATION
    # wait for the confirm button to load, and then click it

    wait = WebDriverWait(driver, 15)

    # 1) Host raíz
    host1 = wait.until(lambda d: d.find_element(By.CSS_SELECTOR, "#LitId"))
    shadow0 = host1.shadow_root

    # 2) Vista create course
    host2 = wait.until(lambda d: shadow0.find_element(By.CSS_SELECTOR, "d2l-create-course-view"))
    shadow1 = host2.shadow_root

    # 3) Botón guardar -> click directo al host
    save_btn = wait.until(lambda d: shadow1.find_element(By.CSS_SELECTOR, "d2l-button[name='saveAndManage']"))

    try:
        save_btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", save_btn)   

    sleep(4)

# The import_course_content method receives the chromedriver, the course object,
# the index of the course and the path of the courses file.
# It imports the course content from the master course, and updates the excel file with the Brightspace ID.
def import_course_content(driver, course, index, courses_path):

    #extraer de la pagina actual el ID del curso recién creado desde la URL (el número después de "ou=") 
    #para luego buscar el curso en la barra de búsqueda y así importar el contenido
    current_url = driver.current_url
    match = re.search(r'ou=(\d+)', current_url)
    if match:
        brightspace_id = match.group(1)
        print(f"\nID de Brightspace obtenido: {brightspace_id}")
    else:
        print("\nNo se encontró el ID de Brightspace en la URL")
        return
   
    #ir a la pagina que importa el contenido del curso (la misma para todos los cursos, 
    #ya que el ID del curso se busca manualmente en la barra de búsqueda)
    driver.get("https://virtual.upb.edu.co/d2l/lms/importExport/import_export.d2l?ou=" + brightspace_id)  #ID de curso genérico para importar contenido, el curso específico se selecciona en la barra de búsqueda

    sleep(2)  # wait for the page to load

    # Manejo de ventanas: guardar el handle de la ventana principal para luego volver a ella después de importar el contenido
    # save the current page for window handling
    mainPage = None
    while not mainPage:
        mainPage = driver.current_window_handle    
    
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
    frame = driver.find_element(By.CSS_SELECTOR, "frame[title='Cuerpo']")

    driver.switch_to.frame(frame)
    sleep(2)

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
    
    #Escribir el ID en Excel sin perder formatos
    update_excel(courses_path, index, brightspace_id)

# The update_excel method receives the file path, the row index and the brightspace_id.
# It updates the column 'Oferta de curso - Id grupo Brightspace' without losing formats and
def update_excel(file_path, row_index, brightspace_id):
    
    """actualiza la columna 'Oferta de curso - Id grupo Brightspace' sin perder formatos y sin afectar otras lineas"""

    # Intenta cargar el archivo Excel
    try:
        wb = load_workbook(file_path)
    except Exception as e:
        logger.exception("Error al abrir el archivo Excel para actualización: %s", file_path)
        print(f"\nError al abrir el archivo: {e}")
        return

    ws = wb.active  # Seleccionar la hoja activa

    # Crear un estilo predeterminado solo si el libro no lo tiene
    if "default" not in wb.named_styles:
        default_style = NamedStyle(name="default")
        wb.add_named_style(default_style)
        
    brightspace_col = 6        # Columna F es la número 6 (1-based)
    excel_row = row_index + 2  # +2 porque pandas indexa desde 0 y Excel tiene encabezado

    try:
        ws.cell(row=excel_row, column=brightspace_col, value=int(brightspace_id))
        print(f"\nID {brightspace_id} agregado a la fila {excel_row} correctamente")
    except Exception as e:
        logger.exception(
            "Error actualizando Brightspace ID en archivo %s, fila %s",
            file_path,
            excel_row,
        )
        print(f"\nError actualizando la celda: {e}")


    # Guardar los cambios en un archivo temporal y luego reemplazar el original
    temp_path = file_path.replace(".xlsx", "_temp.xlsx")
    wb.save(temp_path)
    wb.close()

    # Reemplazar el archivo original con el nuevo
    os.replace(temp_path, file_path)
   
    
# The new_experience method receives the chromedriver and a row of the courses dataframe.
# It checks if the course is a QM course, and if it is, it updates the
def new_experience(driver, row, courses_path, unidades_lookup):
    # Buscar el maestro en el lookup de Unidades precargado
    maestro_value = str(row["Maestro"]).strip()
    name_value = unidades_lookup.get(maestro_value, "").strip()

    if name_value:
        qm_status = "SI" if "QM" in name_value else "NO"

        print(f"\nMaestro: {maestro_value} | Nombre: {name_value} | QM: {qm_status}")

        # Cargar el archivo Excel sin perder formatos
        book = load_workbook(courses_path)
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
        book.save(courses_path)
        book.close()

        print(f"\nRuta de plantilla actualizada en el excel con: {ruta_plantilla}")

        # Recargar el DataFrame después de la modificación con openpyxl
        df_actualizado = pd.read_excel(courses_path, dtype={"Oferta de curso - Id grupo Brightspace": str})
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
                book = load_workbook(courses_path)
                sheet = book.active
                sheet.cell(row=row_index, column=9, value="SI")  # Columna "Se actualizo use lesson experience"
                book.save(courses_path)
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
                book = load_workbook(courses_path)
                sheet = book.active
                sheet.cell(row=row_index, column=10, value="SI")  # Columna "Se tuvo que actualizar content template path"
                book.save(courses_path)
                book.close()

                print("\nActualización de Content Template Path completada, se actualiza el excel con SI.")

            except Exception as e:
                logger.exception(
                    "Error al procesar el curso '%s' en new_experience.",
                    row.get("Nombre", "N/A"),
                )
                print(f"\nError al procesar el curso {row['Nombre']}: {e}")

        else:
            # Si es "NO", actualizar las columnas en Excel con "NO"
            book = load_workbook(courses_path)
            sheet = book.active
            sheet.cell(row=row_index, column=9, value="NO")  # Columna "Se actualizo use lesson experience"
            sheet.cell(row=row_index, column=10, value="NO")  # Columna "Se tuvo que actualizar content template path"
            book.save(courses_path)
            book.close()

            print(f"\nCurso {row['Nombre']}: No requiere actualización de experiencia. Se actualiza el Excel con 'NO'.")
    else:
        print(f"\n Maestro {maestro_value} no encontrado en el lookup de Unidades. Se omite la actualización de experiencia.")
            
# The duplicate_course method receives the chromedriver, a row of the courses dataframe, 
# the index of the course, the path of the courses file, the unidades_lookup and the banner_lookup.
def duplicate_course(driver, course, index, courses_path, unidades_lookup, banner_lookup):
    # access the course management page
    driver.get(
        "https://virtual.upb.edu.co/d2l/platformTools/Courses/6606/createCourse"
    )

    # STEP 1
    select_course_template(driver, course)

    # STEP 2
    select_course_semestre(driver, course)

    # STEP 3
    course_offering_details(driver, course, banner_lookup)

    # STEP 4-
    confirm_course_creation(driver)

    # STEP 5-
    import_course_content(driver, course, index, courses_path)
    
    # STEP 6-
    new_experience(driver, course, courses_path, unidades_lookup)

    return
