# -*- coding: utf-8 -*-
"""
@author: JEMU

MODULO AUXILIAR PARA LOS PROGRAMAS QUE USAN SELENIUM
E INTRODUCEN LOS DATOS EN EL SISTEMA ATRAVÉS DE UNA WEB.
"""
from selenium import webdriver #pip install selenium REQUIRED LIBRARY ASSUMED TO BE INSTALLED
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
import time
from PyQt5.QtWidgets import QLabel, QLineEdit, QPushButton, QVBoxLayout, QDialog, QMessageBox
def _show_warning(message):
    QMessageBox.warning(None, 'Warning', message)
def _Get_login_credentials():
    root = QDialog()
    root.setGeometry(100, 100, 300, 150)
    root.setWindowTitle('Sign In')

    email = QLineEdit()
    password = QLineEdit()
    password.setEchoMode(QLineEdit.Password)

    def login_clicked():
        root.accept()

    def cancel_clicked():
        root.reject()

    email.returnPressed.connect(login_clicked)
    password.returnPressed.connect(login_clicked)

    layout = QVBoxLayout()
    email_label = QLabel('Email Address:')
    layout.addWidget(email_label)
    layout.addWidget(email)
    password_label = QLabel('Password:')
    layout.addWidget(password_label)
    layout.addWidget(password)
    
    login_button = QPushButton('Login')
    login_button.clicked.connect(login_clicked)
    layout.addWidget(login_button)
    
    cancel_button = QPushButton('Cancel')
    cancel_button.clicked.connect(cancel_clicked)
    layout.addWidget(cancel_button)

    root.setLayout(layout)
    root.show()
    root.raise_()
    root.activateWindow()
    
    if root.exec_() == QDialog.Accepted:
        return email.text(), password.text()
    else:
        return None, None
def _logged():
    try:
        driver.find_element(By.CLASS_NAME, "UserProfileLayout---avatar_initials_text")
        print("User is logged in.")
        return True
    except NoSuchElementException:
        mail, pw = _Get_login_credentials()
        if mail is None and pw is None:
            print("Login cancelled by user.")
            return True  # Exit the loop if login is cancelled
        
        search_box = driver.find_element(By.ID, 'un')
        search_box.send_keys(mail)
        search_box = driver.find_element(By.ID, "pw")
        search_box.send_keys(pw)
        search_box.submit()

        try:
            driver.find_element(By.ID, 'login-error')  # Adjust this to your specific login error element
            _show_warning("Incorrect email or password. Please try again.")
        except NoSuchElementException:
            pass
        
        return False

def _login_loop():
    logged_in = False
    while not logged_in:
        logged_in = _logged()


#Input box aux
def _Input_text(title):

    # Root window
    root = QDialog()
    root.setGeometry(100, 100, 350, 100)
    root.setWindowTitle(title)

    txt = QLineEdit()

    def ok_clicked():
        # Callback when the OK button is clicked or Enter is pressed
        root.accept()

    # Connect the returnPressed signal to the ok_clicked function
    txt.returnPressed.connect(ok_clicked)

    # Layout setup
    layout = QVBoxLayout()

    # Input text entry
    layout.addWidget(txt)

    # OK button
    ok_button = QPushButton('OK')
    ok_button.clicked.connect(ok_clicked)
    layout.addWidget(ok_button)

    root.setLayout(layout)
    root.show()
    root.raise_()
    root.activateWindow()

    root.exec_()

    return txt.text()

# Connect to existing instance
def _Conn_open_driver():
    options = Options()
    #options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    driver = webdriver.Chrome(options=options)
    return driver

# Navigate the web 'functions'
def _Menu_bttn():
    Menu_bttn = driver.find_element(By.XPATH, "//div[@class='NavigationLayout---header']//button[contains(@data-testid,'VirtualNavigationH')]")
    Menu_bttn.click()
    time.sleep(1)
def _Clients_bttn():
    #need to call the '_Menu_bttn function'
    _Menu_bttn()
    search_bttn = driver.find_element(By.XPATH, "//li[@title='Gestión Clientes']/button[@class='VirtualNavigationMenuTab_SIDEBAR---nav_tab_clickable_container']")
    search_bttn.click()
    time.sleep(1)
def _Draft_Client_bttn():
    #need to call the '_Menu_bttn function'
    _Menu_bttn()
    search_bttn = driver.find_element(By.XPATH, "//li[@title='Prealta']/a[@data-testid='navigationMenuTab-link']")
    search_bttn.click()
    time.sleep(1)
def _New_Client_bttn():
    #need to call the '_Clients_bttn function'
    _Clients_bttn()
    search_bttn = driver.find_element(By.XPATH, "//li[@title='Nueva alta unificada']/a[@data-testid='navigationMenuTab-link']")
    search_bttn.click()
    time.sleep(1)
def _Search_Client_bttn():
    #need to call the '_Clients_bttn function'
    _Clients_bttn()
    search_bttn = driver.find_element(By.XPATH, "//li[@title='Búsqueda Clientes']/a[@data-testid='navigationMenuTab-link']")
    search_bttn.click()
    time.sleep(1)
def _Search_text():
    #need to call the 'Input_text function'
    input_field=driver.find_element(By.XPATH, "//input[@type='text']")
    id_label=input_field.get_attribute("id")
    label=driver.find_element(By.XPATH, f"//label[@for='{id_label}']")
    txt=_Input_text(label.text)
    input_field.send_keys(txt)
    search_btt=driver.find_element(By.XPATH, "//div[@class='ContentLayout---content_layout']//button[@type='button']")
    search_btt.click()
    time.sleep(1)
def _Search(search_value):
    input_field=driver.find_element(By.XPATH, "//input[@type='text']")
    input_field.send_keys(search_value)
    search_btt=driver.find_element(By.XPATH, "//div[@class='ContentLayout---content_layout']//button[@type='button']")
    search_btt.click()
    time.sleep(1)
def _Confirmation_bttn():
    Confirm_bttn = driver.find_element(By.XPATH, "//div[@class='ContentLayout---content_layout']//button[@id='confirmation']")
    Confirm_bttn.click()
    time.sleep(1)
def Abrir_Deuda_Web(value):
    _login_loop()
    _Search_Client_bttn()
    _Search(value)
    state=driver.find_element(By.XPATH,"//input[@id='client_status']")
    state.send_keys("V")
    state.send_keys(Keys.RETURN)
    time.sleep(1)
    comment=driver.find_element(By.XPATH, "\\input[@id='client_comment']")
    comment.send_keys("CUENTA ABIERTA POR PAGO")       
    comment.send_keys(Keys.RETURN)
    time.sleep(1)
    _Confirmation_bttn()
def Cerrar_Deuda_Web(value):
    _login_loop()
    _Search_Client_bttn()
    _Search(value)
    state=driver.find_element(By.XPATH,"//input[@id='client_status']")
    state.send_keys("V")
    state.send_keys(Keys.RETURN)
    time.sleep(1)
    comment=driver.find_element(By.XPATH, "\\input[@id='client_comment']")
    comment.send_keys("CUENTA CERRADA POR DEUDA")       
    comment.send_keys(Keys.RETURN)
    time.sleep(1)
    _Confirmation_bttn()   
driver = _Conn_open_driver()
driver.switch_to.new_window()
driver.get('https://logistalibrosdev.appiancloud.com/suite/sites/alta-unificada/page/prealtas')

