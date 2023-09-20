from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import os
from pathlib import Path





class WhatsAppAutomator:
    def __init__(self):
        self.BASE_DIR = Path(__file__).resolve().parent
        self.driver = webdriver.Chrome()
        self.wait = WebDriverWait(self.driver, 500)
        self.list_celular = []


    def getter_files(self):
        carta_compromiso_pdf = os.path.join(self.BASE_DIR, "CARTA_COMPROMISO.pdf")
        detalle_beneficiario_pdf = os.path.join(self.BASE_DIR, "detalle_beneficiario.pdf")
        filename_excel = os.path.join(self.BASE_DIR, "envio_wsp.xlsx")
        df = pd.read_excel(filename_excel, header=0)
        documents = [carta_compromiso_pdf, detalle_beneficiario_pdf]
        return df, documents

    def get_dates(self):
        DIA = input("Dia?")
        FECHA = input("Fecha?")
        return (DIA, FECHA)
    
    def open_wsp(self):
        driver = self.driver
        driver.get("https://web.whatsapp.com/")

    def send_messages(self, df, documents, DIA, FECHA):
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
            try:
                time.sleep(2)
                self.driver.find_element(By.XPATH, not_found)
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
            input_message.send_keys(MESSAGE + Keys.ENTER)
            time.sleep(2)
            return False

        def send_document(documents):
            # MANDAS UN DOCUMENTO
            for PDF in documents:
                self.driver.find_element(By.CSS_SELECTOR, "span[data-icon='attach-menu-plus']").click()
                self.driver.find_element(By.CSS_SELECTOR,"input[type='file']").send_keys(PDF)
                time.sleep(3)
                self.driver.find_element(By.CSS_SELECTOR,"span[data-icon='send']").click()

                time.sleep(2)


        for index, row in df.iterrows():
            cel = row['celular']
            OB = row["OB"]
            codigo = row["codigos"]

            br = (Keys.SHIFT)+(Keys.ENTER)+(Keys.SHIFT) 
            MESSAGE = f"*PROGRAMA DE REFUERZO ALIMENTARIO DE BAR-PROVINCIA-MUNI Y CONCEJO*{br}Nos comunicamos para enviarles la confirmación del día y horario asignados para la entrega de alimentos, *ES UNA ENTREGA EVNETUAL* que se hace en Carriego 360:{br}✅ *DIA: {DIA} {FECHA}*{br}✅ *HORA: 14:00*{br}✅ *CÓDIGO: {codigo}*{br}⚠️ *IMPORTANTE* ⚠️{br}✅ *RESPETAR EL HORARIO ASIGNADO*{br}✅ *ENVIARNOS NOMBRE Y DNI DE QUIEN RETIRA{br}COMPLETAR LOS SIGUIENTES DOCUMENTOS ADJUNTADOS Y TRAERLOS IMPRESOS EN EL PRÓXIMO RETIRO DE ALIMENTOS.*"
            
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
                    send_document(documents)
                    self.list_celular.append(cel)
                    



            

    def trigger_automatization(self):
        df,documents = self.getter_files()
        print("Programa de automatización de envios de mensajes")
        print("RECORDATORIOS:\n-Haber guardado excel antes de ingresar al programa\n-Tener agendados los contactos\n-Para mejor funcionamiento no hacer otra actividad en la computadora que demande internet\n-Chequear con el celular si el primer mensaje enviado corresponde a la primera persona en el excel.")
        DIA, FECHA = self.get_dates()
        self.open_wsp()
        self.send_messages(df, documents, DIA, FECHA)
        
        input("Press enter to quit...")
        self.driver.quit()

if __name__ == "__main__":
    automator = WhatsAppAutomator()
    automator.trigger_automatization()





