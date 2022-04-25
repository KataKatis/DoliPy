from time import sleep
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from doli_exceptions import DesignationNonRemplie


class Devis:  # remplissage devis
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

        # open excel (.xlsx) file : template.xlsx
        self.wb = load_workbook(self.file)
        self.ws = self.wb['Devis']

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

    def spec_row_type(self, row):
        row_values = tuple(item.value for item in row)  # get line's values
        if all(x is None for x in row_values):  # check if line is empty --> call self.empty_line function
            self.text_line("   ")
            return True
        elif str(row[0].value).startswith(("[T1]", "[T2]", "[T3]", "[T4]", "[T5]", "[T6]", "[T7]", "[T8]", "[T9]")):  # check if line is title
            self.title(row_values[0])
            return True
        elif str(row[0].value).startswith(("[ST1]", "[ST2]", "[ST3]", "[ST4]", "[ST5]", "[ST6]", "[ST7]", "[ST8]", "[ST9]")):  # check if line is subtotal
            self.subtotal(row[0].value)
            return True
        elif all(x is None for x in row_values[1:]):
            self.text_line(row_values[0])
            return True
        return False

    def text_line(self, text):
        self.driver.execute_script("arguments[0].click();", self.driver.find_element(By.XPATH, "//div[@class='tabsAction']/div[3]/a"))  # click on button "Ajouter une ligne de texte" which execute JS script
        sleep(0.1)
        text_area = self.driver.find_element(By.XPATH, "//div[@id='cke_sub-total-title']/div/div/iframe")
        self.driver.switch_to.frame(text_area)
        self.driver.find_element(By.XPATH, "/html/body").send_keys(text)
        self.driver.switch_to.default_content()
        self.driver.find_element(By.XPATH, "//button[text()='Ok']").click()

    def title(self, title):
        title_lvl = title[2]  # get title level
        title = str(title[4:]).lstrip()  # remove [Tn] and remove if necessary space before text
        self.driver.execute_script("arguments[0].click();", self.driver.find_element(By.ID, "add_title_line"))  # click on button "Ajouter un titre" which execute JS script
        self.driver.find_element(By.ID, "sub-total-title").send_keys(title)
        Select(self.driver.find_element(By.NAME, "subtotal_line_level")).select_by_value(title_lvl)
        self.driver.find_element(By.XPATH, "//button[text()='Ok']").click()

    def subtotal(self, subtotal):
        subtotal_lvl = subtotal[3]  # get subtotal level
        subtotal = str(subtotal[5:]).lstrip()  # remove [Tn] and remove if necessary space before text
        self.driver.execute_script("arguments[0].click();", self.driver.find_element(By.ID, "add_total_line"))  # click on button "Ajouter un sous-total" which execute JS script
        self.driver.find_element(By.ID, "sub-total-title").send_keys(subtotal)
        Select(self.driver.find_element(By.NAME, "subtotal_line_level")).select_by_value(subtotal_lvl)
        self.driver.find_element(By.XPATH, "//button[text()='Ok']").click()

    def unite(self, user_unit):
        user_unit = str(user_unit).lower().strip()  # unit in excel doc
        correct_unite = ("Kg", "Mètre", "Mètre carré", "m³", "dm³", "Pièce", "Heure", "Jour", "Vide", "Ensemble", "Mètre linéaire")  # unit allowed in Dolibarr
        unite = (
            ("kilogramme", "kilogrammes", "kg"),
            ("mètre", "mètres", "m"),
            ("m²", "mètre carré", "mètres carré"),
            ("mètre cube", "mètres cube", "m³"),
            ("décimètre cube", "décimètres cube", "dm³"),
            ("pièce", "u", "unité"),
            ("heure", "heures", "h"),
            ("jour", "jours", "j"),
            ("vide", "", "none"),
            ("ens", "ensemble"),
            ("mètre linéaire", "metre lineaire", "ml")
        )
        for index, unit_type in enumerate(unite):
            if user_unit in unit_type:
                return correct_unite[index]

    def fill_field(self, type_option, designation, TVA, HT, quantite, unite, reduction):
        Select(self.driver.find_element(By.ID, "select_type")).select_by_visible_text(type_option)  # Ligne libre de type
        sleep(0.2)
        # description / designation
        text_area = self.driver.find_element(By.XPATH, "//div[@id='cke_1_contents']/iframe")
        self.driver.switch_to.frame(text_area)
        self.driver.find_element(By.XPATH, "/html/body").send_keys(designation)
        self.driver.switch_to.default_content()
        sleep(0.2)
        Select(self.driver.find_element(By.ID, "tva_tx")).select_by_value(TVA)  # TVA
        sleep(0.2)
        self.driver.find_element(By.ID, "price_ht").send_keys(HT)  # HT
        sleep(0.2)
        self.driver.find_element(By.ID, "qty").clear()  # clear quantity field because "1" by default
        self.driver.find_element(By.ID, "qty").send_keys(quantite)  # quantity
        sleep(0.2)
        Select(self.driver.find_element(By.ID, "units")).select_by_visible_text(unite)  # unite
        sleep(0.2)
        self.driver.find_element(By.ID, "remise_percent").clear()  # clear reduction field because "0" by default
        self.driver.find_element(By.ID, "remise_percent").send_keys(reduction)  # reduction
        sleep(0.2)
        self.driver.find_element(By.ID, "addline").send_keys(Keys.RETURN)  # submit

    def excel(self):
        iteration = self.start

        for row in tuple(self.ws.rows)[self.start:self.end]:
            sleep(0.7)
            # check if line is empty --> call self.empty_line function
            if self.spec_row_type(row):
                iteration += 1
                sleep(0.5)
                continue
            iteration += 1

            # fetch input
            type_option = "Service"
            designation = row[0].value
            if designation is None:
                raise DesignationNonRemplie
            # if/else for making default value
            TVA = "20" if row[1].value is None else (f"{int(row[1].value * 100)}" if row[1].value * 100 == int(row[1].value * 100) else f"{row[1].value * 100}")  # TVA int/float
            HT = "0" if row[2].value is None else row[2].value
            quantite = "0" if row[3].value is None else row[3].value
            unite = self.unite(row[4].value)
            reduction = "0" if row[5].value is None else row[5].value * 100

            self.fill_field(type_option, designation, TVA, HT, quantite, unite, reduction)
