import helpers
import pandas as pd
from getpass import getpass


def main():
    # get the user credentials
    usr = input("Ingrese su usuario: ")
    pwd = getpass("Ingrese su contraseña: ")
    dfa = input("Clave 2FA: ")

    # get the file path
    #path = input("Ingrese la ruta al archivo CSV que contiene la información de los cursos. Puede escribir la ruta o arrastrar el archivo hasta esta ventana: ")
    #path = path.replace('"', "")
    path="courses.xlsx"

    # create the chrome driver
    driver = helpers.create_chrome_driver()

    # login to BS
    helpers.brightspace_login(driver, usr, pwd, dfa)

    try:
        courses = pd.read_excel(path, sheet_name=0)
        cursosFlag = helpers.check_courses_file(courses)
        if not cursosFlag:
            return
    except:
        print(
            "Hubo un error leyendo el archivo "
            + path
            + ". Revise sus entradas e intente de nuevo."
        )
        return

    for index, row in courses.iterrows():
        percentage = ((index + 1) / len(courses)) * 100
        print(
            "Duplicando curso",
            index + 1,
            "de",
            len(courses),
            "(" + str(round(percentage,2)) + "%)",
        )
        helpers.duplicate_course(driver, row)

    print("Terminando la ejecución del programa...")
    driver.close()


if __name__ == "__main__":
    main()
