import os
import helpers
import pandas as pd
from getpass import getpass

def main():

    # Ruta fija del archivo
    path = "courses.xlsx"

    # Verificar si el archivo existe antes de leerlo
    if not os.path.exists(path):
        print(f"Error: No se encontró el archivo '{path}'. Verifique la ruta y el nombre.")
        return

    # Crear el driver de Chrome
    driver = helpers.create_chrome_driver()

    # Iniciar sesión en Brightspace
    helpers.brightspace_login(driver)

    try:
        # Leer el archivo Excel asegurando que los nombres de las columnas no tengan espacios extra
        courses = pd.read_excel(path, sheet_name=0)
        courses.columns = courses.columns.str.strip()  # Limpiar espacios en nombres de columnas

        # Verificar si el archivo está vacío
        if courses.empty:
            print("Error: El archivo courses.xlsx está vacío.")
            return

        # Validar el formato del archivo
        cursosFlag = helpers.check_courses_file(courses)
        if not cursosFlag:
            print("Error: Formato incorrecto del archivo de cursos. Verifique las columnas.")
            return

    except pd.errors.EmptyDataError:
        print("Error: El archivo courses.xlsx no contiene datos.")
        return
    except pd.errors.ParserError:
        print("Error: No se pudo leer el archivo courses.xlsx. Revise el formato.")
        return
    except Exception as e:
        print(f"Hubo un error inesperado al leer el archivo: {e}")
        return

    # Iterar sobre cada fila del DataFrame
    for index, row in courses.iterrows():
        percentage = ((index + 1) / len(courses)) * 100
        print(
            "\nDuplicando curso",
            index + 1,
            "de",
            len(courses),
            "(" + str(round(percentage, 2)) + "%)",
        )
        helpers.duplicate_course(driver, row)

    print("\nTerminando la ejecución del programa...")
    driver.close()


if __name__ == "__main__":
    main()