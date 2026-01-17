from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from time import sleep

import pandas as pd

"""The brightspace_login method receives three arguments: the chromedriver, the user and the password of the Virtual Campus admin.
# If the login is succesful, it return True, otherwise, it returns False.
"""

def brightspace_login(driver, usr, pwd, dfa):
    try:
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
        username.send_keys(usr)
        #password.send_keys(pwd)
        #password.send_keys("2/2*2Hijosmemsam")
        password.send_keys("Sza.2751176")
        password.send_keys(Keys.RETURN)
        sleep(1)

        #2FA
        driver.get("https://virtual.upb.edu.co/d2l/lp/auth/twofactorauthentication/TwoFactorCodeEntry.d2l")
        
        l2fa = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.ID, "z_d"))
        )
        l2fa.send_keys(dfa)
        l2fa.send_keys(Keys.RETURN)

        return True
    except:
        print(
            "Hubo un error al tratar de iniciar sesión. Revise sus credenciales e intente de nuevo.❌"
        )

        return False


"""The create_chrome_driver receives no arguments. It creates the Chrome Driver with options to run on headless mode.
It returns the created driver.
"""


def create_chrome_driver():
    # service object
    s = Service("..\Chrome\chromedriver.exe")

    # chrome options to run on headless mode and no logging on the terminal
    chrome_options = Options()
    chrome_options.add_argument("--disable-extensions")
    #chrome_options.add_argument("--headless")
    chrome_options.add_argument("--log-level=3")

    # instantiate the webdriver
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

    if len(columns) != 5:
        return False

    return (
        "Maestro" in columns
        and "Nombre" in columns
        and "Codigo" in columns
        and "Plantilla" in columns
        and "Semestre" in columns
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


def import_course_content(driver, course):
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
    sleep(1)

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
    sleep(1)

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
    sleep(1)

    #find the frame containing the search bar for the course offering
    #frame = driver.find_element(By.CSS_SELECTOR, "#PopupWindow > frame:nth-child(3)")   --- sin UserWay
    #frame = driver.find_element(By.XPATH, "/html/frameset/frame[2]")
    #frame = driver.find_element(By.CSS_SELECTOR, "#PopupWindow > frame:nth-child(10)")  -- cambio 2023/11/15
    frame = driver.find_element(By.CSS_SELECTOR, "frame[title='Cuerpo']")

    driver.switch_to.frame(frame)
    sleep(1)

    #searchInput = driver.find_element(By.XPATH, '//input[@id="z_b"]')
    searchInput = driver.find_element(By.CSS_SELECTOR, "#z_b")
    
    searchInput.send_keys(str(course["Maestro"]))
    searchInput.send_keys(Keys.RETURN)
  
    driver.find_element(
        By.CSS_SELECTOR,
        "#yui-rec0 > td.d_dg_col_d_selection.yui-dt0-col-d_selection.yui-dt-col-d_selection.yui-dt-first > div > span > input",
    ).click()

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


def duplicate_course(driver, course):
    # access the course duplication page
    driver.get(
        "https://virtual.upb.edu.co/d2l/tools/courseCreate/courseCreateType.asp?ou=6606"
    )

    # STEP 1,
    select_course_template(driver, course)

    # STEP 2,
    course_offering_details(driver, course)

    # STEP 3,
    confirm_course_creation(driver)

    # STEP 4,
    import_course_content(driver, course)

    return
