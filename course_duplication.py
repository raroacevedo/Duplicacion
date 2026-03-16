import json
import logging
import os
from datetime import datetime
from pathlib import Path

import pandas as pd

import helpers

# Constantes para la configuración de archivos y claves requeridas en el JSON de configuración
CONFIG_FILE = "external_files.json"
REQUIRED_FILE_KEYS = ("path", "unidades_path", "BDBANNER")
REQUIRED_CONFIG_KEYS = REQUIRED_FILE_KEYS + ("DIRLOGS",)

# Configuración de logging con manejo de archivos y consola
def setup_logging(logs_dir):
    logs_path = Path(logs_dir)
    logs_path.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    log_name = f"log_{timestamp}.log"
    log_path = logs_path / log_name

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(formatter)

    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)

    logging.info("Inicio de ejecución. Archivo de log: %s", log_path)
    return str(log_path)

# Resolver rutas relativas al directorio del script o rutas absolutas según corresponda
def resolve_path(path_value):
    path_obj = Path(path_value)
    if path_obj.is_absolute():
        return path_obj
    return (Path(__file__).resolve().parent / path_obj).resolve()

# Cargar y validar el archivo de configuración JSON que contiene las rutas de los archivos externos
def load_external_files_config(config_filename=CONFIG_FILE):
    config_path = resolve_path(config_filename)
    if not config_path.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo de configuración requerido: {config_path}"
        )

    with open(config_path, "r", encoding="utf-8") as json_file:
        config = json.load(json_file)

    missing_keys = [key for key in REQUIRED_CONFIG_KEYS if key not in config]
    if missing_keys:
        raise KeyError(
            "Faltan claves obligatorias en el JSON de configuración: "
            + ", ".join(missing_keys)
        )

    resolved_config = {}
    for key in REQUIRED_CONFIG_KEYS:
        value = str(config[key]).strip()
        if not value:
            raise ValueError(f"La clave '{key}' no tiene una ruta válida en el JSON.")
        resolved_config[key] = str(resolve_path(value))

    return resolved_config

# Validar que los archivos obligatorios existen antes de iniciar el proceso
def validate_required_files(files_config):
    missing_files = []

    for key in REQUIRED_FILE_KEYS:
        file_path = files_config[key]
        if not os.path.exists(file_path):
            logging.error("Archivo obligatorio no encontrado (%s): %s", key, file_path)
            print(f"Error: No se encontró el archivo obligatorio '{file_path}'.")
            missing_files.append(file_path)

    return len(missing_files) == 0


# Cargar los datos de unidades desde el archivo Excel y crear un diccionario de búsqueda 
# para facilitar el acceso a los nombres de las unidades por su código
def load_unidades_lookup(unidades_path):
    unidades_df = pd.read_excel(unidades_path, dtype={"Code": str, "Name": str})
    required_cols = {"Code", "Name"}
    missing_cols = required_cols - set(unidades_df.columns)
    if missing_cols:
        raise ValueError(
            "El archivo de unidades no tiene las columnas requeridas: "
            + ", ".join(sorted(missing_cols))
        )

    clean_df = unidades_df.loc[:, ["Code", "Name"]].dropna(subset=["Code"]).copy() # Eliminar filas donde "Code" es NaN

    return dict(zip(clean_df["Code"], clean_df["Name"]))

# Cargar los datos de Banner desde el archivo Excel y crear un diccionario de búsqueda
# que permita acceder a las fechas de inicio y fin de curso por combinación de lista cruzada y periodo
def load_banner_lookup(banner_path):
    # Asegurando que se lea la primera hoja del archivo de Banner y que las columnas de interés se lean como texto para evitar problemas de formato
    banner_df = pd.read_excel(banner_path, sheet_name=0, dtype={"LISTA_CRUZADA": str, "PERIODO": str}) 
    required_cols = {"LISTA_CRUZADA", "PERIODO", "FECHA_INICIO_CURSO", "FECHA_FIN_CURSO"}
    missing_cols = required_cols - set(banner_df.columns)
    if missing_cols:
        raise ValueError(
            "El archivo Banner no tiene las columnas requeridas: "
            + ", ".join(sorted(missing_cols))
        )
    
    # Limpiar y normalizar los datos de Banner para asegurar que las claves de búsqueda sean consistentes
    clean_df = banner_df.loc[:, ["LISTA_CRUZADA", "PERIODO", "FECHA_INICIO_CURSO", "FECHA_FIN_CURSO"]].copy()

    banner_lookup = {}
    for row in clean_df.itertuples(index=False):

        nrc=row.LISTA_CRUZADA
        periodo=row.PERIODO
        if not nrc or not periodo:
            continue

        key = (nrc, periodo)
        # Si existen duplicados de clave en Banner, se conserva el primero válido.
        if key not in banner_lookup:
            banner_lookup[key] = (row.FECHA_INICIO_CURSO, row.FECHA_FIN_CURSO)

    return banner_lookup

# Función principal que orquesta la ejecución del programa, incluyendo la carga de configuraciones, validación de archivos,
# precarga de datos, manejo del driver de Chrome, iteración sobre los cursos y manejo de excepc
def main():
    try:
        files_config = load_external_files_config()

        setup_logging(files_config["DIRLOGS"])

        if not validate_required_files(files_config):
            logging.error("Validación de archivos obligatorios fallida. Proceso detenido.")
            return

        courses_path = files_config["path"]
        unidades_path = files_config["unidades_path"]
        banner_path = files_config["BDBANNER"]

        print("Cargando datos de unidades de organización...")
        unidades_lookup = load_unidades_lookup(unidades_path)
        print("Cargando datos de Banner...")
        banner_lookup = load_banner_lookup(banner_path)

        logging.info(
            "Datos precargados: %s maestros en unidades y %s registros de Banner.",
            len(unidades_lookup),
            len(banner_lookup),
        )

        driver = None
        try:
            # Crear el driver de Chrome
            driver = helpers.create_chrome_driver()

            # Iniciar sesión en Brightspace
            if not helpers.brightspace_login(driver):
                logging.error("No fue posible iniciar sesión en Brightspace.")
                return

            # Leer el archivo Excel asegurando que los nombres de columnas no tengan espacios extra
            courses = pd.read_excel(courses_path, sheet_name=0)
            courses.columns = courses.columns.str.strip()  # Limpiar espacios en nombres de columnas

            # Verificar si el archivo está vacío
            if courses.empty:
                logging.error("El archivo de cursos está vacío: %s", courses_path)
                print("Error: El archivo de cursos está vacío.")
                return

            # Validar el formato del archivo
            cursos_flag = helpers.check_courses_file(courses)
            if not cursos_flag:
                logging.error("Formato incorrecto del archivo de cursos: %s", courses_path)
                print("Error: Formato incorrecto del archivo de cursos. Verifique las columnas.")
                return

            total_courses = len(courses)

            # Iterar sobre cada fila del DataFrame
            for index, row in courses.iterrows():
                percentage = ((index + 1) / len(courses)) * 100
                logging.info(
                    "Iteración %s/%s (%.2f%%) - Curso: %s",
                    index + 1,
                    total_courses,
                    percentage,
                    row.get("Nombre", "N/A"),
                )

                try:
                    helpers.duplicate_course(driver, row, index, courses_path, unidades_lookup, banner_lookup)
                    logging.info(
                        "Iteración %s/%s completada correctamente.",
                        index + 1,
                        total_courses,
                    )
                except Exception:
                    logging.exception(
                        "Error en iteración %s/%s. Fila Excel %s (Nombre: %s).",
                        index + 1,
                        total_courses,
                        index + 2,
                        row.get("Nombre", "N/A"),
                    )

            logging.info("Terminando la ejecución del programa...")

        except pd.errors.EmptyDataError:
            logging.exception("El archivo de cursos no contiene datos: %s", courses_path)
            print("Error: El archivo de cursos no contiene datos.")
        except pd.errors.ParserError:
            logging.exception("No se pudo leer el archivo de cursos: %s", courses_path)
            print("Error: No se pudo leer el archivo de cursos. Revise el formato.")
        except Exception:
            logging.exception("Error inesperado durante la ejecución principal.")
            print("Hubo un error inesperado. Revise el archivo de log para más detalle.")
        finally:
            if driver is not None:
                driver.close()
                logging.info("Navegador cerrado.")

    except Exception:
        logging.basicConfig(level=logging.ERROR, format="%(asctime)s [%(levelname)s] %(message)s")
        logging.exception("Error de configuración inicial.")
        print("Error de configuración inicial. Revise el archivo de log para más detalle.")


if __name__ == "__main__":
    main()
