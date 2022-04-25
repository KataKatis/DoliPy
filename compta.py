from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from doli_exceptions import NumCompteNonRempli

class Compta:
    def __init__(self, url, username, password, start, end):
        self.url = url
        self.username = username
        self.password = password
        self.start = start
        self.end = end
        self.file = "template.xlsx"

        self.driver = webdriver.Chrome()
        self.driver.get(self.url)
        self.driver.maximize_window()

        self.connect()
        self.excel()

        input()

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

    def fill_field(self, num_compte, compte_auxiliaire, libelle_compte_auxiliaire, libelle_operation, debit, credit):
        Select(self.driver.find_element(By.ID, "accountingaccount_number")).select_by_value(str(num_compte))
        if compte_auxiliaire is not None : self.driver.find_element(By.NAME, "subledger_account").send_keys(compte_auxiliaire)
        if libelle_compte_auxiliaire is not None : self.driver.find_element(By.NAME, "subledger_label").send_keys(libelle_compte_auxiliaire)
        self.driver.find_element(By.NAME, "label_operation").clear()  # not in if
        if libelle_operation is not None : self.driver.find_element(By.NAME, "label_operation").send_keys(libelle_operation)
        if debit is not None : self.driver.find_element(By.NAME, "debit").send_keys(debit)
        if credit is not None : self.driver.find_element(By.NAME, "credit").send_keys(credit)
        self.driver.find_element(By.NAME, "save").send_keys(Keys.RETURN)

    def excel(self):
        wb = load_workbook(self.file)
        ws = wb['Compta']

        for row in tuple(ws.rows)[self.start:self.end]:
            num_compte = row[0].value
            if num_compte is None:
                raise NumCompteNonRempli
            compte_auxiliaire = row[1].value
            libelle_compte_auxiliaire = row[2].value
            libelle_operation = row[3].value
            debit = row[4].value
            credit = row[5].value

            self.fill_field(num_compte, compte_auxiliaire, libelle_compte_auxiliaire, libelle_operation, debit, credit)
