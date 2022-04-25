from openpyxl import load_workbook
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from time import sleep


class NoteDeFrais:  # remplissage note de frais
    def __init__(self, url, username, password, start, end):
        self.url = url
        self.username = username
        self.password = password
        self.start = start
        self.end = end
        self.file = "template.xlsx"

        # open chrome
        self.driver = webdriver.Chrome()
        self.driver.get(self.url)
        self.driver.maximize_window()

        self.connect()
        self.excel()

        # closing chromedriver.exe (if not, processus will stay in background, to delete it : Ctrl + Alt + Suppr > Task Manager)
        self.driver.quit()


    def connect(self):  # account connection
        username_area = self.driver.find_element(By.ID, "username")
        password_area = self.driver.find_element(By.ID, "password")
        username_area.clear()
        password_area.clear()
        username_area.send_keys(self.username)
        password_area.send_keys(self.password)
        password_area.send_keys(Keys.RETURN)
        self.driver.get(self.url)

    def fill_field(self, date, expense_type, description, TVA, TTC, quantite):
        self.driver.find_element(By.ID, "date").send_keys(date)  # date
        sleep(0.2)
        Select(self.driver.find_element(By.ID, "fk_c_type_fees")).select_by_visible_text(expense_type)  # type
        sleep(0.2)
        self.driver.find_element(By.XPATH, "//textarea").send_keys(description)  # description
        sleep(0.2)
        Select(self.driver.find_element(By.ID, "vatrate")).select_by_value(TVA)  # TVA
        sleep(0.2)
        self.driver.find_element(By.ID, "value_unit").send_keys(TTC)  # TTC
        sleep(0.2)
        self.driver.find_element(By.XPATH, "//input[@name='qty']").clear()  # clear quantity field because "1" by default
        self.driver.find_element(By.XPATH, "//input[@name='qty']").send_keys(quantite)  # quantity
        sleep(0.2)
        self.driver.find_element(By.XPATH, "//input[@value='Ajouter']").click()  # submit


    def excel(self):
        # open excel (.xlsx) file : template.xlsx
        wb = load_workbook(self.file)
        ws = wb['NDF']

        for row in tuple(ws.rows)[self.start:self.end]:  # for each line from self.start to self.end
            date = str(datetime.strptime(str(row[0].value), "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y"))  # convert date format from excel to date format for dolibarr
            expense_type = row[1].value
            description = "" if row[2].value is None else row[2].value
            # if TVA an integer --> int type else --> float type (because select TVA with value in HTML)
            TVA = "20" if row[3].value is None else (f"{int(row[3].value * 100)}" if row[3].value * 100 == int(row[3].value * 100) else f"{row[3].value * 100}")
            TTC = "0" if row[4].value is None else row[4].value
            quantite = "0" if row[5].value is None else row[5].value

            self.fill_field(date, expense_type, description, TVA, TTC, quantite)
