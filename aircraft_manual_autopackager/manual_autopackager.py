import sys
import os
import random
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QLabel, QWidget, QLineEdit, QFileDialog, QFrame, QHBoxLayout
from PyQt5.QtGui import QFont
from PyQt5 import QtCore
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import threading


class PDFScraper(QMainWindow):
    def __init__(self):
        super().__init__()
        self.driver = None
        self.load_last_session()  # Load last session values
        self.init_ui()
        self.loop_flag = False

    def init_ui(self):
        self.setWindowTitle("Scrape AirbusWorld")
        self.setGeometry(100, 100, 400, 250)

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)

        # Add title label
        title_label = QLabel("AirbusWorld Scrapper TBS2",
                             font=QFont("Arial", 16, QFont.Bold))
        title_label.setAlignment(QtCore.Qt.AlignCenter)
        layout.addWidget(title_label)

        self.username_edit = QLineEdit()
        self.username_edit.setPlaceholderText("Enter Username")
        self.username_edit.setText(self.username)  # Set last session username
        layout.addWidget(self.username_edit)

        self.password_edit = QLineEdit()
        self.password_edit.setPlaceholderText("Enter Password")
        self.password_edit.setEchoMode(QLineEdit.Password)
        self.password_edit.setText(self.password)  # Set last session password
        layout.addWidget(self.password_edit)

        self.customization_edit = QLineEdit()
        self.customization_edit.setPlaceholderText("Enter Customization")
        # Set last session customization
        self.customization_edit.setText(self.customization)
        layout.addWidget(self.customization_edit)

        self.msn_edit = QLineEdit()
        self.msn_edit.setPlaceholderText("Enter MSN")
        self.msn_edit.setText(self.MSN)  # Set last session MSN
        layout.addWidget(self.msn_edit)

        # Create a wrapper widget for the label, button, and separator
        file_wrapper_widget = QWidget()
        file_layout = QHBoxLayout(file_wrapper_widget)
        # Separator line
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        file_layout.addWidget(separator, 1)  # Set ratio to 1
        # File path label
        self.file_label = QLabel("No file selected.")
        file_layout.addWidget(self.file_label)
        # File browser button
        self.file_button = QPushButton("Select MPD/AMM List File")
        self.file_button.clicked.connect(self.select_excel_file)
        file_layout.addWidget(self.file_button)
        # Wrap layout
        layout.addWidget(file_wrapper_widget)

        self.start_button = QPushButton("Start Login")
        self.start_button.clicked.connect(self.mulai_login)
        layout.addWidget(self.start_button)

        # Scrape by MPD button
        self.scrape_mpd_button = QPushButton("Scrape into Taskcard")
        self.scrape_mpd_button.clicked.connect(self.start_scraping_mpd)
        layout.addWidget(self.scrape_mpd_button)

        # Stop Scraping button
        self.stop_scraping_button = QPushButton("Stop Scraping Taskcard")
        self.stop_scraping_button.clicked.connect(self.stop_scraping_mpd)
        layout.addWidget(self.stop_scraping_button)

        # Download Report
        self.download_report_button = QPushButton("Download Report Scraping")
        self.download_report_button.clicked.connect(self.download_report)
        layout.addWidget(self.download_report_button)

        # Scrape by MPD button
        self.scrape_amm_button = QPushButton("Scrape into Non Taskcard")
        self.scrape_amm_button.clicked.connect(self.start_scraping_amm)
        layout.addWidget(self.scrape_amm_button)

        # Stop Scraping button
        self.stop_scraping_amm_button = QPushButton(
            "Stop Scraping Non Taskcard")
        self.stop_scraping_amm_button.clicked.connect(self.stop_scraping_amm)
        layout.addWidget(self.stop_scraping_amm_button)

        self.status_label = QLabel("")
        layout.addWidget(self.status_label)

    def load_last_session(self):
        # Check if the settings file exists
        if os.path.isfile("last_session.txt"):
            with open("last_session.txt", "r") as file:
                lines = file.readlines()
                self.username = lines[0].strip()
                self.password = lines[1].strip()
                self.customization = lines[2].strip()
                self.MSN = lines[3].strip()
        else:
            self.username = ""
            self.password = ""
            self.customization = ""
            self.MSN = ""

    def save_last_session(self):
        with open("last_session.txt", "w") as file:
            file.write(self.username + "\n")
            file.write(self.password + "\n")
            file.write(self.customization + "\n")
            file.write(self.MSN)

    def mulai_login(self):
        self.username = self.username_edit.text()
        self.password = self.password_edit.text()
        self.customization = self.customization_edit.text()
        self.MSN = self.msn_edit.text()
        self.save_last_session()  # Save current session values

        waktu_tunggu = 20
        self.link_text = f'https://w3.airbus.com/1T40/search/text?wc=actype:A330;customization:{self.customization};doctype:AMM;tailNumber:F{self.MSN}'

        chrome_options = Options()
        chrome_options.add_argument("--headless")
        # self.driver = webdriver.Chrome("chromedriver.exe", options=chrome_options)
        self.driver = webdriver.Chrome()

        # Login phase
        self.driver.get("https://w3.airbus.com/")
        WebDriverWait(self.driver, waktu_tunggu).until(
            EC.element_to_be_clickable((By.LINK_TEXT, "click here"))).click()
        WebDriverWait(self.driver, waktu_tunggu).until(EC.presence_of_element_located(
            (By.NAME, "USER"))).send_keys(self.username)
        WebDriverWait(self.driver, waktu_tunggu).until(EC.presence_of_element_located(
            (By.NAME, "PASSWORD"))).send_keys(self.password)
        time.sleep(1)
        WebDriverWait(self.driver, waktu_tunggu).until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "#loginContent > form > input.btn-primary.btn-submit"))).click()
        self.driver.get(self.link_text)
        self.status_label.setText("Login Successful!")  # Update status label

    def select_excel_file(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx *.xls);;All Files (*)", options=options)
        if file_name:
            print("Selected file:", file_name)
            self.file_label.setText(file_name)
            try:
                self.refs = pd.read_excel(file_name, sheet_name=0)[
                    'ref'].to_list()
            except:
                print("Error create list")

    def download_report(self):
        _report = pd.DataFrame({'Reference': pd.Series(self.refs),
                                'Terpaket': pd.Series(self.list_qt)})
        output_file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Report in excel", "", "Excel Files (*.xlsx);;All Files (*)")
        if output_file_path:
            with open(output_file_path, "wb") as output_file:
                _report.to_excel(output_file)
            self.status_label.setText(
                f"Merged PDF saved to: {output_file_path}")

    def start_scraping_mpd(self):
        self.loop_flag = True
        self.stop_scraping_button.setEnabled(True)
        self.scrape_mpd_button.setEnabled(False)
        threading.Thread(target=self.scrape_by_mpd).start()

    def stop_scraping_mpd(self):
        self.loop_flag = False
        self.stop_scraping_button.setEnabled(False)
        self.scrape_mpd_button.setEnabled(True)

    def scrape_by_mpd(self):
        # buat def dulu biar simple
        self.list_qt = []

        def masukkan_paket_konfirm():
            _btn = WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button[ng-click='$mdOpenMenu($event)']")))
            time.sleep(0.7)
            _btn.click()
            _btn = WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "a[ng-if='vm.isJobCardActive']")))
            time.sleep(0.5)
            _btn.click()
            _btn = WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button[ng-click='callback()']")))
            time.sleep(0.5)
            _btn.click()
            time.sleep(0.5)

        def masukkan_paket():
            _btn = WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button[ng-click='$mdOpenMenu($event)']")))
            time.sleep(0.7)
            _btn.click()
            _btn = WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "a[ng-if='vm.isJobCardActive']")))
            time.sleep(0.5)
            _btn.click()
            time.sleep(0.5)

        # Clear jobcard package dlu
        WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "button[ng-click='navbarCtrl.jobCard()']"))).click()
        WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "button[ng-click='confirmClearJobCardBasket($event)']"))).click()
        WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "button[ng-click='callback()']"))).click()
        WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "button[ng-click='callback()']"))).click()

        awal = True
        # mulai loop
        for ref in self.refs:
            if not self.loop_flag:
                break  # Stop the loop if loop_flag is False

            len_ref = len(ref)
            qt_terpaket = 0

            WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "input[ng-model='SearchBarCtrl.searchText']"))).clear()
            WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "input[ng-model='SearchBarCtrl.searchText']"))).send_keys(f'{ref}*')
            WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "input[ng-model='SearchBarCtrl.searchText']"))).send_keys(Keys.ENTER)

            try:
                result_elements = WebDriverWait(self.driver, 5).until(EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, "div[ng-repeat='item in STextCtrl.content']")))
            except:
                result_elements = None
                print('Reference Not found')

            if result_elements:
                if len(result_elements) > 1 and len_ref == 9:
                    for result in result_elements:
                        try:
                            result.click()
                        except:
                            print("error ngklik cuy, coba klik lagi")
                            result.click()
                        if awal:
                            try:
                                masukkan_paket_konfirm()
                                WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable(
                                    (By.CSS_SELECTOR, "div[class='toast-message']"))).click()
                                qt_terpaket += 1
                                awal = False
                            except:
                                try:
                                    masukkan_paket_konfirm()
                                    qt_terpaket += 1
                                except:
                                    print("Error packaging")
                        else:
                            try:
                                masukkan_paket()
                                WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable(
                                    (By.CSS_SELECTOR, "div[class='toast-message']"))).click()
                                qt_terpaket += 1
                            except:
                                try:
                                    masukkan_paket()
                                    qt_terpaket += 1
                                except:
                                    print("Error packaging")
                else:
                    try:
                        result_elements[0].click()
                    except:
                        print("error ngklik cuy, coba klik lagi")
                        result_elements[0].click()
                    if awal:
                        try:
                            masukkan_paket_konfirm()
                            WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable(
                                (By.CSS_SELECTOR, "div[class='toast-message']"))).click()
                            qt_terpaket += 1
                            awal = False
                        except:
                            try:
                                masukkan_paket_konfirm()
                                qt_terpaket += 1
                            except:
                                print("Error packaging")
                    else:
                        try:
                            masukkan_paket()
                            WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable(
                                (By.CSS_SELECTOR, "div[class='toast-message']"))).click()
                            qt_terpaket += 1
                        except:
                            try:
                                masukkan_paket()
                                qt_terpaket += 1
                            except:
                                print("Error packaging")
                self.list_qt.append(qt_terpaket)
                time.sleep(random.randrange(2, 3))

    def scrape_amm(self):
        # buat def dulu biar simple
        def masukkan_paket_konfirm():
            _btn = WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button[aria-label='Print']")))
            time.sleep(0.7)
            _btn.click()
            _btn = WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "a[ng-click='vm.addTaskToPdfBasket($event)']")))
            time.sleep(0.5)
            _btn.click()
            _btn = WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button[ng-click='callback()']")))
            time.sleep(0.5)
            _btn.click()
            time.sleep(0.5)

        def masukkan_paket():
            _btn = WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button[aria-label='Print']")))
            time.sleep(0.7)
            _btn.click()
            _btn = WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "a[ng-click='vm.addTaskToPdfBasket($event)']")))
            time.sleep(0.5)
            _btn.click()
            time.sleep(0.5)

        awal = True
        # mulai loop
        for ref in self.refs:
            if not self.loop_flag:
                break  # Stop the loop if loop_flag is False

            WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "input[ng-model='SearchBarCtrl.searchText']"))).clear()
            WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "input[ng-model='SearchBarCtrl.searchText']"))).send_keys(f'{ref}*')
            WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "input[ng-model='SearchBarCtrl.searchText']"))).send_keys(Keys.ENTER)

            try:
                result_elements = WebDriverWait(self.driver, 5).until(EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, "div[ng-repeat='item in STextCtrl.content']")))
            except:
                result_elements = None
                print('Reference Not found')

            if result_elements:
                try:
                    result_elements[0].click()
                except:
                    print("error ngklik cuy, coba klik lagi")
                    result_elements[0].click()
                if awal:
                    try:
                        masukkan_paket_konfirm()
                        WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable(
                            (By.CSS_SELECTOR, "div[class='toast-message']"))).click()
                        awal = False
                    except:
                        try:
                            masukkan_paket_konfirm()
                        except:
                            print("Error packaging")
                else:
                    try:
                        masukkan_paket()
                        WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable(
                            (By.CSS_SELECTOR, "div[class='toast-message']"))).click()
                    except:
                        try:
                            masukkan_paket()
                        except:
                            print("Error packaging")

                time.sleep(random.randrange(2, 3))

    def start_scraping_amm(self):
        self.loop_flag = True
        self.stop_scraping_amm_button.setEnabled(True)
        self.scrape_amm_button.setEnabled(False)
        threading.Thread(target=self.scrape_amm).start()

    def stop_scraping_amm(self):
        self.loop_flag = False
        self.stop_scraping_amm_button.setEnabled(False)
        self.scrape_amm_button.setEnabled(True)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFScraper()
    window.show()
    sys.exit(app.exec_())
