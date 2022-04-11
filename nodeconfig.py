# import module
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
import time
import pandas as pd

options = Options()
#options.headless = True
options.add_experimental_option('excludeSwitches', ['enable-logging'])

# read "macAddress" excel file
df = pd.read_excel('macAddress.xlsx')
macAddress = df["MAC"].tolist()

def wifi(x): 
    # function to establish a new connection
    def createNewConnection(name, SSID, password):
        config = """<?xml version=\"1.0\"?>
    <WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
        <name>"""+name+"""</name>
        <SSIDConfig>
            <SSID>
                <name>"""+SSID+"""</name>
            </SSID>
        </SSIDConfig>
        <connectionType>ESS</connectionType>
        <connectionMode>auto</connectionMode>
        <MSM>
            <security>
                <authEncryption>
                    <authentication>WPA2PSK</authentication>
                    <encryption>AES</encryption>
                    <useOneX>false</useOneX>
                </authEncryption>
                <sharedKey>
                    <keyType>passPhrase</keyType>
                    <protected>false</protected>
                    <keyMaterial>"""+password+"""</keyMaterial>
                </sharedKey>
            </security>
        </MSM>
    </WLANProfile>"""
        command = "netsh wlan add profile filename=\""+name+".xml\""+" interface=Wi-Fi"
        with open(name+".xml", 'w') as file:
            file.write(config)
        os.system(command)
    
    # function to connect to a network   
    def connect(name, SSID):
        command = "netsh wlan connect name=\""+name+"\" ssid=\""+SSID+"\" interface=Wi-Fi"
        os.system(command)
    
    name = "SBnode_"+ x.upper()
    password = "12345678"
    
    # establish new connection
    createNewConnection(name, name, password)
    
    # connect to the wifi network
    connect(name, name)

def cellular():
    #open networking menu
    url = "http://192.168.8.1/cgi-bin/luci/admin/network/status"
    driver = webdriver.Chrome(options=options)
    driver.get(url)
    #gateway login
    passfield = driver.find_element(By.NAME, "luci_password")
    loginbutton = driver.find_element(By.XPATH, """//*[@id="maincontent"]/form/div[2]/input[1]""")
    passfield.send_keys("admin")
    loginbutton.click()

    
    #select connectivity
    select_element = driver.find_element(By.ID,'wanid')
    wandropdown = Select(select_element)
    saveapply = driver.find_element(By.XPATH,"""//*[@id="maincontent"]/form/fieldset[2]/table/tbody/tr[2]/td[2]/input""")

    wandropdown.select_by_visible_text("Cellular")  
    saveapply.click()
    #accept allert
    alert = driver.switch_to.alert
    alert.accept()
    time.sleep(3)

    driver.quit()

for x in macAddress:
    # print(x)
    wifi(x)
    time.sleep(5)
    cellular()
    time.sleep(5)
