# Use silceq environment
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains

import pandas as pd
import time
import os

from dotenv import load_dotenv
load_dotenv()

# Configuration
db_file = "3A.xlsx"
silceq_user = os.getenv("SILCEQ_USER")
silceq_pass = os.getenv("SILCEQ_PASS")
silceq_grado = os.getenv("SILCEQ_GRADO")
silceq_grupo = os.getenv("SILCEQ_GRUPO")

# Load database
db = pd.read_excel(db_file)
print(db)

# db.iloc[st,su] for student st and subject su

# Login to silceq
URI_SILCEQ = "http://silceq.usebeq.edu.mx/sce/Entrada.php?usuacla=&usuapass="
URI_SILCEQ2 = "http://silceq.usebeq.edu.mx/sce/MANTTO_GEN.PHP?manto=Reg_Apo_Req&usua=" + silceq_user
chromeOptions = Options()
chromeOptions.headless = True
drs = webdriver.Chrome(options=chromeOptions)
drs.get(URI_SILCEQ)
delay = 3  # seconds
myElem = WebDriverWait(drs, delay).until(
    EC.presence_of_element_located((By.ID, "usuario")))
input_user = drs.find_element_by_name("usuario")
input_password = drs.find_element_by_name("contrasena")
input_user.send_keys(silceq_user)
input_password.send_keys(silceq_pass)
input_user.submit()
# Get students lists for grade and group selected
drs.get(URI_SILCEQ2)
drs.find_element_by_xpath("//input[@id='gra']").send_keys(silceq_grado)
drs.find_element_by_xpath("//input[@id='gru']").send_keys(silceq_grupo)
drs.find_element_by_xpath(
    "//div[@id='CollapsiblePanel1']/div/table[2]/tbody/tr/td/center/input[1]").click()


# Obtain table of students and open a tab for each one
# Create subjects and route for each one of them
subjects = [
    {'cat': 'CBA',
     'catx': 'CURRICULA BASICA                                            ',
             'id': '100',
             'idx': 'Lenguaje y Comunicación'},
    {'cat': 'CBA',
     'catx': 'CURRICULA BASICA                                            ',
             'id': '101',
             'idx': 'Pensamiento Matemático'},
    {'cat': 'CBA',
     'catx': 'CURRICULA BASICA                                            ',
             'id': '102',
             'idx': 'Exploración y Comprensión del Mundo Natural y Social	'},
    {'cat': 'ARS',
     'catx': 'AREAS                                                       ',
             'id': '103',
             'idx': 'Artes'},
    #            {'cat':'ARS',
    #            'catx':'AREAS                                                       ',
    #            'id':'113',
    #            'idx':'Educación Socioemocional'},
    {'cat': 'ARS',
     'catx': 'AREAS                                                       ',
             'id': '105',
             'idx': 'Educación Física'}]
bn_set = drs.find_elements_by_xpath(
    "//input[@type='image' and @src='../images/nuevo24x24.png']")
for sid in range(len(subjects)):

    drs.switch_to.window(drs.window_handles[0])
    # drs.find_element_by_xpath("//input[@name='m_tm_clave']").send_keys(s.get('cat'))
    # Check if db and number of students is the same
    print(len(bn_set))
    print(len(db))
    assert len(bn_set) == len(db)
    s = subjects[sid]
    for k in reversed(range(len(bn_set))):
        bn = bn_set[k]
        ActionChains(drs) \
            .key_down(Keys.CONTROL) \
            .click(bn) \
            .key_up(Keys.CONTROL) \
            .perform()
    # Insert data for the student on each tab
    for st in range(len(db)):
        tabid = st+1
        drs.switch_to.window(drs.window_handles[1])
        drs.execute_script(
            "document.getElementsByName('m_destipmat')[0].removeAttribute('readonly')")
        drs.execute_script(
            "document.getElementsByName('m_desmat')[0].removeAttribute('readonly')")
        drs.find_element_by_xpath(
            "//input[@name='m_tm_clave']").send_keys(s.get('cat'))
        # drs.find_element_by_xpath("//input[@name='m_destipmat']").send_keys(s.get('catx'))
        drs.find_element_by_xpath(
            "//input[@name='m_ma_clave']").send_keys(s.get('id'))
        # drs.find_element_by_xpath("//input[@name='m_desmat']").send_keys(s.get('idx'))
        drs.find_element_by_xpath(
            "//textarea[@name='m_aa_mensaje']").send_keys(db.iloc[st, sid+1])
        # time.sleep(400)
        drs.find_element_by_xpath(
            "//input[@type='image' and @src='../images/Save32x32.png']").click()
        time.sleep(0.5)
        drs.close()
