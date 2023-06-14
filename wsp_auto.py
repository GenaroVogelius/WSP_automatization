from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time


# ? browser
driver=webdriver.Chrome()

# ? url to open
driver.get("https://web.whatsapp.com/")

# ? how long wait will be in seconds.
wait=WebDriverWait(driver,500)


filename = r'C:\Users\Usuario\Downloads\envio_wsp.xlsx'

df = pd.read_excel(filename, header=0)

DIA = input("dia?")
FECHA = input("fecha?")
list_celular = []

# bucle start getting data from excel:
for index, row in df.iterrows():
    codigo = row['CÓDIGO']
    cel = row['CELULAR']
    OB = row["NOMBRE OB"]
    
    
    
    br = (Keys.SHIFT)+(Keys.ENTER)+(Keys.SHIFT) 
    message = f"*PROGRAMA DE REFUERZO ALIMENTARIO DE BAR-PROVINCIA-MUNI Y CONCEJO*{br}Nos comunicamos para enviarles la confirmación del día y horario asignados para la entrega de alimentos, que se hacen en Carriego 360:{br} ✅ *DIA: {DIA} {FECHA}*{br} ✅ *HORA: 14:00*{br} ✅ *CÓDIGO: {codigo}*{br} ⚠️ *IMPORTANTE* ⚠️{br} ✅ *RESPETAR EL HORARIO ASIGNADO*{br} ✅ *ENVIARNOS NOMBRE Y DNI DE QUIEN RETIRA*"
    

    # ? catch the search
    search_tab = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="side"]/div[1]/div/div/div[2]')))
    search_tab.click()

    # ? write in the search tab the number to send message
    input_tab = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="side"]/div[1]/div/div/div[2]/    div/    div[1]  ')))
    
    input_tab.send_keys(cel)
    
    # ? click the first user
    time.sleep(3)
    
    
    input_tab.send_keys(Keys.ARROW_DOWN)
    input_tab.send_keys(Keys.ENTER)
    
    not_found = '//span[text()="No se encontró ningún chat, contacto ni mensaje."]'
    try:
        driver.find_element(By.XPATH, not_found)
        print(f"No se ha encontrado la organización barrial *{OB}* con el numero *{cel}*, contiene el código *{codigo}*")
        input_tab.send_keys(Keys.CONTROL + "a")
        input_tab.send_keys(Keys.DELETE)
        continue
    except:
        pass

    
    # ? catch the input for writting message
    time.sleep(2)
    input_user = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#main > footer > div._2lSWV._3cjY2.copyable-area > div > span:nth-child(2) > div > div._1VZX7 > div._3Uu1_ > div > div.to2l77zo.gfz4du6o.ag5g9lrv.bze30y65.kao4egtt')))
    
    if cel in list_celular:
        input_user.send_keys(f"También tenes el siguiente código para retirar {codigo}" + Keys.ENTER)
    else:
        input_user.send_keys(message + Keys.ENTER)
        list_celular.append(cel)
    
    input_tab.send_keys(Keys.CONTROL + "a")
    input_tab.send_keys(Keys.DELETE)
    

    # ? CHEQUER
    # is_okey = input("is okey?")
    # if is_okey == "y":
    #     input_user.send_keys(Keys.ENTER)
    # else:
    #     input_user.send_keys(Keys.CONTROL + "a")
    #     input_user.send_keys(Keys.DELETE)
    #     input_tab.send_keys(Keys.CONTROL + "a")
    #     input_tab.send_keys(Keys.DELETE)
    #     time.sleep(1)
    #     continue
    # print("se mando el mensaje")



input("Press enter to quit...")
driver.quit()


