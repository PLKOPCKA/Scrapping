from msedge.selenium_tools import Edge, EdgeOptions
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from Password import fico_pass
from SQL_Run_Proc import *
import time
from datetime import datetime
import os

start_time = time.time()

bands = [0.3, 0.45, 0.6, 0.7, 0.78, 0.84, 0.9, 0.95, 1.01]
cuts = [95, 95, 90, 90, 90, 90, 90, 90, 90]
max_solve = 14400
min_gap = 1
scenario = 'NewTesting_FixThe3Lanes'
model = 'CUAL'

options = EdgeOptions()
options.use_chromium = True
options.binary_location = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
driver = Edge(executable_path=r"E:\00_SQL Queries\Kamil\msedgedriver.exe", options=options)

driver.get('http://scs-prd-as02:8860/insight/login.jsp')
driver.maximize_window()

login = driver.find_element_by_id('inputEmail')
login.send_keys('PLKOPCKA')

password = driver.find_element_by_id('inputPassword')
password.send_keys(fico_pass)

connect = driver.find_element_by_id('loginBtn')
connect.click()
time.sleep(5)

app_button = driver.find_element_by_xpath("//div[contains(@class, 'project-clickable-area')]")
app_button.click()

time.sleep(3)
app_button = driver.find_element_by_xpath("//ul[contains(@class, 'shelf-item-list col-xs-11 scenario-manager-launcher ui-sortable')]")
app_button.click()


def open_folder(name, double_click=False):
    done = False
    i = 0
    time.sleep(3)
    for table in driver.find_elements_by_xpath("//div[contains(@class, 'scenario-manager-items-list-container')]"):
        for row in table.find_elements_by_xpath("//div[contains(@class, 'tabulator-table')]"):
            for col in row.find_elements_by_xpath("//div[contains(@class, 'tabulator-row tabulator-selectable tabulator-row-odd')]"):
                if col.text == name:
                    i += 1
                    if i == 2:
                        if double_click:
                            chains = ActionChains(driver)
                            chains.double_click(col).perform()
                        else:
                            col.click()
                        done = True
                        break
            if done:
                break
        if done:
            break
    return


def navigate_to_frame(name):
    time.sleep(5)
    for el2 in WebDriverWait(driver, 120).until(EC.visibility_of_any_elements_located((By.XPATH,"//a[contains(@class, 'tab-item view-select-item')]"))):
        if el2.text == name:
            el2.click()
            break


def go_to_app():
    for el in WebDriverWait(driver, 120).until(EC.visibility_of_any_elements_located((By.XPATH, "//ul[contains(@class, 'nav nav-pills pull-left')]/li"))):
        if el.text == 'APP':
            el.click()
            break


def click_job():
    for el in WebDriverWait(driver, 120).until(EC.visibility_of_any_elements_located((By.XPATH, "//ul[contains(@class, 'nav nav-pills pull-left')]/li"))):
        if el.text == 'JOBS':
            el.click()
            break


def select_right_panel(name, iframe=True):
    time.sleep(5)
    if iframe:
        WebDriverWait(driver, 120).until(EC.frame_to_be_available_and_switch_to_it("view-iframe"))
    for el in WebDriverWait(driver, 120).until(EC.visibility_of_any_elements_located((By.XPATH, "//ul[contains(@class, 'sidebar compact-sidebar list-group')]/li/a"))):
        if el.text == name:
            el.click()
            break


def go_to_job(operation, max_wait=3600, check=5, get_time=False):
    time.sleep(5)
    click_job()

    i = 0
    t = 0
    done = False
    try:
        while i <= max_wait:
            i += check
            t += check
            if t >= 200:
                click_job()
                t = 0
            time.sleep(check)
            j = 0
            for row in WebDriverWait(driver, 120).until(EC.visibility_of_any_elements_located((By.XPATH, "//tr[contains(@data-id, '8df3e616-4224-4aa9-8cac-f9dda59ef203')]"))):
                j = 0
                model = ''
                process = ''
                status = ''
                for col in row.find_elements_by_tag_name('td'):
                    j += 1
                    if j == 2: model = col.text
                    if j == 3: process = col.text
                    if j == 5: status = col.text
                    if j == 10 and col.text != '': exec_time = col.text
                    if model == 'CUAL' and process == operation and status == 'Completed':
                        done = True
                        break
                if done:
                    break
            if done:
                break
    except:
        pass

    go_to_app()

    if get_time:
        return exec_time


def click_button(name, loc=1):
    time.sleep(5)
    i = 0
    for el in WebDriverWait(driver, 120).until(EC.visibility_of_any_elements_located((By.XPATH, "//button[contains(@class, 'btn btn-primary')]"))):
        if el.text == name:
            i += 1
            if i == loc:
                el.click()
                break

    if name == 'DOWNLOAD RESULTS':
        select_right_panel('Results Export', False)
    driver.switch_to.default_content()


def lp_relaxation(tick=False):
    time.sleep(10)
    i = 0
    for el in WebDriverWait(driver, 120).until(EC.visibility_of_any_elements_located((By.XPATH, "//div[contains(@class, 'checkbox')]/label"))):
        i += 1
        if i == 4:
            if tick:
                if not el.find_element_by_tag_name("input").is_selected():
                    el.click()
            else:
                if el.find_element_by_tag_name("input").is_selected():
                    el.click()
            break


def unclick_messages():
    time.sleep(3)
    try:
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it("view-iframe"))
        for el in WebDriverWait(driver, 10).until(EC.visibility_of_any_elements_located((By.XPATH, "//div[contains(@id, 'message')]/div"))):
            el.click()
    except:
        pass
    driver.switch_to.default_content()


open_folder('Network_Design_Models')
open_folder('CustomerAllocation')
open_folder('Transformation')

# select the CUAL model from the folder
time.sleep(3)
done = False
for row in driver.find_elements_by_xpath("//div[contains(@class, 'tabulator-row tabulator-selectable tabulator-row-even')]"):
    for col in row.find_elements_by_xpath("//div[contains(@class, 'tabulator-cell')]"):
        if col.text == model:
            chains = ActionChains(driver)
            chains.double_click(col).perform()
            done = True
            break
    if done:
        break

app_button.click()

navigate_to_frame('Optimization')
select_right_panel('Transportation')
time.sleep(10)
click_button('DOWNLOAD RESULTS')

for k, band in enumerate(bands):
    if k > 1:
        time.sleep(30)
        cut = cuts[k]
        if k > 0:
            update_lanes_values(scenario, bands[k-1])

        execute_sql_proc(cut, band, k)

        navigate_to_frame('Setup')
        select_right_panel('Mass Data Import')
        click_button('IMPORT DATA FROM DATABASE', 2)

        go_to_job('RUN_ALL_QUERIES', 600, 5)

        navigate_to_frame('Optimization')
        unclick_messages()
        try:
            select_right_panel('Optimization')
        except:
            time.sleep(30)
            driver.refresh()
            unclick_messages()
            select_right_panel('Optimization')

        # # Adjust Solver settings max time solve and min gap value
        # z = 0
        # p = 0
        # for panel in WebDriverWait(driver, 5).until(EC.visibility_of_any_elements_located((By.XPATH, "//div[contains(@class, 'checkbox')]"))):
        #     p += 1
        #     if p == 1:
        #         for el in panel.find_elements_by_xpath("//input[contains(@class, 'form-control')]"):
        #             z += 1
        #             if z == 1:
        #                 el.clear()
        #                 el.send_keys(max_solve)
        #             if z == 2:
        #                 el.clear()
        #                 el.send_keys(min_gap)

        lp_relaxation(False)
        click_button('OPTIMIZE')
        exec_time = go_to_job('OPTIMIZE', max_solve, 60, True)
        navigate_to_frame('Optimization')
        time.sleep(20)
        navigate_to_frame('Optimization')
        unclick_messages()
        try:
            select_right_panel('Optimization')
        except:
            time.sleep(20)
            driver.refresh()
            unclick_messages()
            select_right_panel('Optimization')

        # collect the data about Objective Function value and gap
        z = 0
        for row in driver.find_elements_by_tag_name('h5'):
            z += 1
            if z == 2:
                objective = row.text
            if z == 3:
                gap = row.text

        driver.switch_to.default_content()

        # datetime object containing current date and time
        now = datetime.now()
        dt_string = now.strftime("%d/%m/%Y %H:%M:%S")

        insert_outputs(scenario, band, dt_string, objective, gap, exec_time)

        time.sleep(5)
        navigate_to_frame('Results')
        driver.refresh()
        time.sleep(30)
        select_right_panel('Results Export')
        click_button('POPULATE RESULTS')
        go_to_job('POPULATE_RESULTS', 600, 3)
        driver.refresh()
        unclick_messages()
        select_right_panel('Transportation')
        time.sleep(10)
        click_button('DOWNLOAD RESULTS')


print("------------------------------------------------------------------")
print("--- %s seconds ---" % (time.time() - start_time))

driver.close()
driver.quit()


latest_file = 'C:/Users/PLKOPCKA/Downloads/Transportation Lane Products Results.csv'
os.remove(latest_file)


