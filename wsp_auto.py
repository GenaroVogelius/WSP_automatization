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
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from io import BytesIO
import win32clipboard
from PIL import Image



class WhatsAppAutomator:
    def __init__(self, message, name_of_excel, images=None, documents=None):


        # ?  si pones esto no te anda
        #  ? chrome_path = ChromeDriverManager().install()

        options = Options()
        options.add_argument('--ignore-certificate-errors')
        options.add_argument("--js-flags=--harmony")
        options.add_argument("--disable-extensions")
        options.add_argument("--enable-automation")
        options.add_argument("--start-maximized")
        options.add_argument("--disable-browser-side-navigation")


        self.BASE_DIR = Path(__file__).resolve().parent
        self.driver = webdriver.Chrome(options=options)
        self.wait = WebDriverWait(self.driver, 500)
        self.short_wait = WebDriverWait(self.driver, 5)

        self.list_celular = []

        
        self.images_names = images
        self.documents_names = documents
        self.message = pyperclip.copy(message)
        self.filename_excel = os.path.join(self.BASE_DIR, name_of_excel)

        self.df, self.documents_list, self.images_list = self.getter_files()


    def getter_files(self):

        images_list = [os.path.join(self.BASE_DIR, image_name) for image_name in self.images_names]
        documents_list = [os.path.join(self.BASE_DIR, document_name) for document_name in self.documents_names]
        df = pd.read_excel(self.filename_excel, header=0)
        return df, documents_list, images_list

    # def get_excel_variables(self):
    #     for index, row in self.df.iterrows():
    #         cel = row['CELULAR']
    #         OB = row["NOMBRE"]



    # private functions 
    def _click_in_search_tab(self):
            search_tab = self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="side"]/div[1]/div/div/div[2]')))
            search_tab.click()

    def _write_in_search_tab(self, cel):
            # ? write in the search tab the number to send message
            self.input_tab = self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="side"]/div[1]/div/div/div[2]/            div/    div[1]  ')))
            self.input_tab.send_keys(cel)


    def _is_a_repeat_number(self, cel):
        if cel in self.list_celular:
            self.input_tab.send_keys(Keys.CONTROL + "a")
            self.input_tab.send_keys(Keys.DELETE)
            return True
        else:
            return False
        
    def _click_first_user(self):
            # click the first user
            time.sleep(3)
            self.input_tab.send_keys(Keys.ARROW_DOWN)
            self.input_tab.send_keys(Keys.ENTER)
        

    def _checker_not_found(self,cel, nombre):
            not_found = '//span[text()="No se encontró ningún chat, contacto ni mensaje."]'
            searching = '//span[text()="Buscando chats, contactos y mensajes…"]'
            
            try:
                time.sleep(2)
                self.driver.find_element(By.XPATH, not_found)
                print(f"No se ha encontrado a {nombre} con el numero {cel}")
                self.input_tab.send_keys(Keys.CONTROL + "a")
                self.input_tab.send_keys(Keys.DELETE)
                return True
            except:
                try:
                    self.driver.find_element(By.XPATH, searching)
                    print(f"No se ha encontrado a {nombre} con el numero {cel}")
                    self.input_tab.send_keys(Keys.CONTROL + "a")
                    self.input_tab.send_keys(Keys.DELETE)
                    return True
                except:
                    return False


    def _write_message(self):
            self.input_message = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#main > footer > div._2lSWV._3cjY2.copyable-area > div > span:nth-child(2) > div > div._1VZX7 > div._3Uu1_ > div > div.to2l77zo.gfz4du6o.ag5g9lrv.bze30y65.kao4egtt')))
            self.input_message.click()
            ActionChains(self.driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
            self.input_message.send_keys(Keys.ENTER)
            time.sleep(2)
            return False


    def _send_images(self):
        for image_path in self.images_list:
            image = Image.open(image_path)
            output = BytesIO()
            image.convert('RGB').save(output, 'BMP')
            data = output.getvalue()[14:]
            output.close()
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            win32clipboard.CloseClipboard()


            # act.key_down(Keys.CONTROL).send_keys("v").key_up(Keys.CONTROL).perform()
            ActionChains(self.driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
            self.input_message.send_keys(Keys.ENTER)
            self.driver.find_element(By.CSS_SELECTOR,"span[data-icon='send']").click()


    def _send_documents(self):
        for pdf in self.documents_list:
        # clipButton = self.wait.until(EC.presence_of_element_located((    By.XPATH,    '//[@id="main"]/footer//*[@data-icon="attach-menu-plus"]/..',)))
                
            self.driver.find_element(By.CSS_SELECTOR, "span[data-icon='attach-menu-plus']").click()
            time.sleep(3)
            self.driver.find_element(By.CSS_SELECTOR,"input[type='file']").send_keys(pdf)
            time.sleep(2)
            self.driver.find_element(By.CSS_SELECTOR,"span[data-icon='send']").click()
            time.sleep(2)



    def start_rows_iteration(self):
        for index, row in self.df.iterrows():
            cel = row['CELULAR']
            nombre = row["NOMBRE"]
            if pd.isna(cel) or pd.isna(nombre):
                 continue
            
            # Remove '-' from the 'CELULAR' column
            self.df.at[index, 'CELULAR'] = str(cel).replace('-', '')


            already_send_a_message = self._is_a_repeat_number(cel)
            if already_send_a_message:
                continue
            else:
                self._click_in_search_tab()
                self._write_in_search_tab(cel)
                self._click_first_user()
                is_not_found =self._checker_not_found(cel, nombre)
                if is_not_found:
                    continue
                else:
                    self._write_message()
                    if self.images_list:
                        self._send_images()
                    if self.documents_list:
                        self._send_documents()
                    self.list_celular.append(cel)



    def open_wsp(self):
        driver = self.driver
        driver.get("https://web.whatsapp.com/")


    def trigger_automatization(self):
        print("Programa de automatización de envios de mensajes")
        print("RECORDATORIOS:\n-Haber guardado excel antes de ingresar al programa\n-Tener agendados los contactos\n-Para mejor funcionamiento no hacer otra actividad en la computadora que demande internet\n-Chequear con el celular si el primer mensaje enviado corresponde a la primera persona en el excel.")
        self.open_wsp()
        self.start_rows_iteration()


        
        input("Presiona enter para cerrar...")
        self.driver.quit()


if __name__ == "__main__":
    images = []
    documents = []


    cantidad_de_cajas = input("Cuantas cajas se entregan? Escribir numero ")
    dia = input("Que dia? ej. hoy, mañana ")
    fecha = input("fecha? ej. 13/11 ")
    message = f"""Hola, como va? Hoy hay *entrega especial de CARNE*, SON {cantidad_de_cajas} CAJAS de aprox 25 kg cada caja. Fue una donación que con mucho trabajo conseguimos 💪. 
Gracias por su constante apoyo. 
FECHA Y HORARIO:  *{dia} {fecha} de 14 a 16 hs.*
            """
    
    name_of_excel= "envio_wsp.xlsx"

    automator = WhatsAppAutomator(message, name_of_excel, images, documents)
    automator.trigger_automatization()





