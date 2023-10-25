from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import os
from pathlib import Path
import pyperclip
from selenium.webdriver import ActionChains
from PIL import Image
import pyperclip

from io import BytesIO
import win32clipboard
from PIL import Image



class WhatsAppAutomator:
    def __init__(self):
        self.BASE_DIR = Path(__file__).resolve().parent
        self.driver = webdriver.Chrome()
        self.wait = WebDriverWait(self.driver, 500)
        self.list_celular = []


    def getter_files(self):
        imagen_path = os.path.join(self.BASE_DIR, "gondolazo.jpg")
        filename_excel = os.path.join(self.BASE_DIR, "excel_filtrado2.xlsx")
        df = pd.read_excel(filename_excel, header=0, dtype={'CELULAR': str})
        documents = [imagen_path]
        return df, documents

    def get_dates(self):
        DIA = input("Dia?")
        FECHA = input("Fecha?")
        return (DIA, FECHA)
    
    def open_wsp(self):
        driver = self.driver
        driver.get("https://web.whatsapp.com/")

    def send_messages(self, df, documents):
        def click_in_search_tab():
            search_tab = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="side"]/div[1]/div/div/div[2]')))
            search_tab.click()

        def write_in_search_tab():
            # ? write in the search tab the number to send message
            input_tab = self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="side"]/div[1]/div/div/div[2]/            div/    div[1]  ')))

            input_tab.send_keys(cel)
            return input_tab

        def click_first_user(input_tab):
            # ? click the first user
            time.sleep(3)
            input_tab.send_keys(Keys.ARROW_DOWN)
            input_tab.send_keys(Keys.ENTER)

        def checker_not_found():
            not_found = '//span[text()="No se encontró ningún chat, contacto ni mensaje."]'
            searching = '//span[text()="Buscando chats, contactos y mensajes…"]'
            
            try:
                time.sleep(2)
                self.driver.find_element(By.XPATH, not_found)
                print(f"No se ha encontrado la organización barrial {OB} con el numero {cel}")
                input_tab.send_keys(Keys.CONTROL + "a")
                input_tab.send_keys(Keys.DELETE)
                return True
            except:
                try:
                    self.driver.find_element(By.XPATH, searching)
                    print(f"No se ha encontrado la organización barrial {OB} con el numero {cel}")
                    input_tab.send_keys(Keys.CONTROL + "a")
                    input_tab.send_keys(Keys.DELETE)
                    return True
                except:
                    return False

        def repeat_number():
            if cel in self.list_celular:
                input_tab.send_keys(Keys.CONTROL + "a")
                input_tab.send_keys(Keys.DELETE)
                return True
            else:
                return False
            
        def write_message():
            # catch the input for writting message
            input_message = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#main > footer > div._2lSWV._3cjY2.copyable-area > div > span:nth-child(2) > div > div._1VZX7 > div._3Uu1_ > div > div.to2l77zo.gfz4du6o.ag5g9lrv.bze30y65.kao4egtt')))
            input_message.click()
            # act.key_down(Keys.CONTROL).send_keys("v").key_up(Keys.CONTROL).perform()
            ActionChains(self.driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
            input_message.send_keys(Keys.ENTER)
            time.sleep(2)
            return False

        def send_document(documents):
            # MANDAS UN DOCUMENTO
            for PDF in documents:
                # clipButton = self.wait.until(EC.presence_of_element_located((    By.XPATH,    '//[@id="main"]/footer//*[@data-icon="attach-menu-plus"]/..',)))
                
                self.driver.find_element(By.CSS_SELECTOR, "span[data-icon='attach-menu-plus']").click()
                time.sleep(3)
                self.driver.find_element(By.CSS_SELECTOR,"input[type='file']").send_keys(PDF)
                time.sleep(2)
                self.driver.find_element(By.CSS_SELECTOR,"span[data-icon='send']").click()

                time.sleep(2)

        def send_images(images):
            for image_path in images:
                image = Image.open(image_path)
                output = BytesIO()
                image.convert('RGB').save(output, 'BMP')
                data = output.getvalue()[14:]
                output.close()
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
                win32clipboard.CloseClipboard()
                time.sleep(2)
                input_message = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#main > footer > div._2lSWV._3cjY2.copyable-area > div > span:nth-child(2) > div > div._1VZX7 > div._3Uu1_ > div > div.to2l77zo.gfz4du6o.ag5g9lrv.bze30y65.kao4egtt')))
                input_message.click()
                # act.key_down(Keys.CONTROL).send_keys("v").key_up(Keys.CONTROL).perform()
                ActionChains(self.driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
                time.sleep(2)
                input_message.send_keys(Keys.ENTER)
                self.driver.find_element(By.CSS_SELECTOR,"span[data-icon='send']").click()
                time.sleep(2)

        for index, row in df.iterrows():

 
            cel = row['CELULAR']
            OB = row["NOMBRE"]
            if pd.isna(cel) or pd.isna(OB):
                 continue
            

            MESSAGE = f"""Hola {OB}, ya falta cada vez menos para nuestra *SUPER ACCIÓN ANUAL*❗

Y nuestros voluntarios, son el motor y la esencia de cada *GONDOLAZO*🤜🏻🤛🏻

Durante el dos días vamos a estar en más de 70 bocas de supermercados para contar quiénes somos, lo que hacemos y buscando donaciones de alimentos para las *más de 300 organizaciones sociales* que lo necesitan de forma URGENTE.

Gondolazo es posible gracias a ustedes, desde hace seis años somos testigos de eso
¿Nos acompañás un año más? 🙏🏻

Podés sumarte *completando el siguiente formulario ¡Y listo!* 📋🙌🏻
Link del formulario: https://forms.gle/TnxrMzJwiFmSySAw9

👉🏻Si quieren sumar a un amigo, amiga, familiar para vivir esta gran experiencia juntos, ¡Ayudanos a difundir!

¡TODOS A INSCRIBIRSE QUE YA LLEGA EL 6TO GONDOLAZO!🛒🤝"""

            MESSAGE = pyperclip.copy(MESSAGE)
            
            already_send_a_message = repeat_number()
            if already_send_a_message:
                continue
            else:
                click_in_search_tab()
                input_tab = write_in_search_tab()
                click_first_user(input_tab)
                is_not_found = checker_not_found()
                if is_not_found:
                    continue
                else:
                    write_message()
                    # send_document(documents)
                    send_images(documents)
                    self.list_celular.append(cel)
                    
            

    def trigger_automatization(self):
        df,documents = self.getter_files()
        print("Programa de automatización de envios de mensajes")
        print("RECORDATORIOS:\n-Haber guardado excel antes de ingresar al programa\n-Tener agendados los contactos\n-Para mejor funcionamiento no hacer otra actividad en la computadora que demande internet\n-Chequear con el celular si el primer mensaje enviado corresponde a la primera persona en el excel.")
        self.open_wsp()
        self.send_messages(df, documents)
        
        input("Press enter to quit...")
        self.driver.quit()

if __name__ == "__main__":
    automator = WhatsAppAutomator()
    automator.trigger_automatization()





