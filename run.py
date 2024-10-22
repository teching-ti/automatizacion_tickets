import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException
from selenium.webdriver.chrome.options import Options
from datetime import datetime
import time

# Configuración del archivo de log
logging.basicConfig(
    filename='tickets_log.txt',  # Nombre del archivo de log
    level=logging.INFO,  # Nivel de log (INFO en este caso)
    format='%(asctime)s -> %(message)s',  # Formato del mensaje (hora y mensaje)
    datefmt='%d-%m-%Y %H:%M:%S'  # Formato de la fecha y hora
)

'''
Automatización para marcar como pendiente o esperando aprobación los tickets que hayan sido creados fuera del horario laboral.
'''

def revisarTicketsFueraHorarioLaboral():
    chrome_options = Options()

    #instancia para instalar el driver cada vez que se realice la ejecución
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    # enlace
    itop_url = "http://192.168.1.253:8081/itop/web/pages/UI.php"
    username = "ti-admin"
    password = "ti@admin2024*-*"
    
    # instancia del driver al enlace
    driver.get(itop_url)
    time.sleep(5)
    driver.find_element(By.ID, "user").send_keys(username)
    driver.find_element(By.ID, "pwd").send_keys(password)
    driver.find_element(By.XPATH, "//input[@type='submit' and @value='Ingresar']").click()

    def buscar_requerimientos():
        driver.find_element(By.ID, "AccordionMenu_RequestManagement").click()
        time.sleep(1)
        driver.find_element(By.LINK_TEXT, "Búsqueda de Requerimientos de Usuario").click()
        time.sleep(1)
        driver.find_element(By.CLASS_NAME, 'sfb_header').click()
        time.sleep(2)

    def buscar_incidentes():
        driver.find_element(By.ID, "AccordionMenu_IncidentManagement").click()
        time.sleep(1)
        driver.find_element(By.LINK_TEXT, "Búsqueda por Incidentes").click()
        time.sleep(1)
        driver.find_element(By.CLASS_NAME, 'sfb_header').click()
        time.sleep(2)

    def revisar_requerimientos():
        max_intentos = 5
        intentos = 0

        while intentos < max_intentos:
            logging.info("Revisando Requerimientos")
            buscar_requerimientos()

            try:
                time.sleep(2)
                tabla = driver.find_element(By.CLASS_NAME, "listResults")
                rows = tabla.find_elements(By.XPATH, "./tbody/tr")[:5]

                for r in rows:
                    id_ticket = r.find_element(By.XPATH, ".//td[1]").text
                    fecha_hora_ticket = r.find_element(By.XPATH, ".//td[5]").text
                    estado_ticket = r.find_element(By.XPATH, ".//td[6]").text
                    
                    ticket_hora = datetime.strptime(fecha_hora_ticket, "%Y-%m-%d %H:%M:%S")
                    dia_de_semana = ticket_hora.weekday()
                    inicio_labores = ticket_hora.replace(hour=8, minute=0, second=0, microsecond=0)
                    fin_labores = ticket_hora.replace(hour=18, minute=0, second=0, microsecond=0)
                    
                    if (ticket_hora < inicio_labores or ticket_hora > fin_labores or dia_de_semana >= 5) and estado_ticket == 'Nuevo':
                        logging.info(f"Ticket {id_ticket} está fuera del horario laboral, creado a las {fecha_hora_ticket}, se procede a colocar como 'Esperando Aprobación'")
                        r.find_element(By.LINK_TEXT, id_ticket).click()
                        time.sleep(3)
                        driver.find_element(By.CLASS_NAME, "actions_menu").click()
                        time.sleep(3)
                        driver.find_element(By.CSS_SELECTOR, "[data-action-id='ev_wait_for_approval']").click()
                        time.sleep(3)
                        driver.find_element(By.XPATH, "//button[@type='submit']/span[text()='Esperando Aprobación']/..").click()
                        time.sleep(3)
                        break
                else:
                    time.sleep(3)
                    break
            except StaleElementReferenceException:
                logging.error("Elemento obsoleto encontrado, reintentando...")
                time.sleep(2)
                intentos += 1
            except NoSuchElementException:
                logging.error("Elemento no encontrado, reintentando...")
                time.sleep(2)
                intentos += 1
            except Exception as e:
                logging.error(f"Error inesperado: {e}")
                break

        if intentos == max_intentos:
            logging.error("Máximo número de reintentos alcanzado. No se pudo completar la operación.")

    def revisar_incidentes():
        max_intentos = 5
        intentos = 0

        while intentos<max_intentos:
            logging.info("Revisando Incidentes")
            # se dirige a la página donde se encuentran los incidentes
            buscar_incidentes()

            try:
                time.sleep(2)
                tabla = driver.find_element(By.CLASS_NAME, "listResults")
                rows = tabla.find_elements(By.XPATH, "./tbody/tr")[:5]

                for r in rows:
                    id_ticket = r.find_element(By.XPATH, ".//td[1]").text
                    fecha_hora_ticket = r.find_element(By.XPATH, ".//td[5]").text
                    estado_ticket = r.find_element(By.XPATH, ".//td[6]").text

                    # se obtiene la fecha en que se ha registrado el ticket
                    ticket_hora = datetime.strptime(fecha_hora_ticket, "%Y-%m-%d %H:%M:%S")
                    dia_de_semana = ticket_hora.weekday()
                    inicio_labores = ticket_hora.replace(hour=8, minute=0, second=0, microsecond=0)
                    fin_labores = ticket_hora.replace(hour=18, minute=00, second=0, microsecond=0)#18

                    # condicion para saber si un ticket ha sido creado fuera del horario laboral
                    if (ticket_hora < inicio_labores or ticket_hora > fin_labores or dia_de_semana >= 5) and estado_ticket == 'Nuevo':
                        logging.info(f"Ticket {id_ticket} está fuera del horario laboral, creado a las {fecha_hora_ticket}, se procede a colocar como 'Pendiente'")
                        r.find_element(By.LINK_TEXT, id_ticket).click()
                        time.sleep(3)
                        driver.find_element(By.CLASS_NAME, "actions_menu").click()
                        time.sleep(3)
                        driver.find_element(By.LINK_TEXT, "Asignar").click()
                        time.sleep(3)
                        driver.find_element(By.XPATH, "//select[@name='attr_team_id']").click()
                        time.sleep(2)
                        driver.find_element(By.XPATH, "//option[@value='40']").click()
                        time.sleep(2)
                        driver.find_element(By.XPATH, "//select[@name='attr_agent_id']").click()
                        time.sleep(2)
                        driver.find_element(By.XPATH, "//option[@value='143']").click()
                        time.sleep(2)
                        driver.find_element(By.XPATH, "//button//span[text()='Asignar']").click()
                        time.sleep(2)
                        driver.find_element(By.CLASS_NAME, "actions_menu").click()
                        time.sleep(2)
                        driver.find_element(By.LINK_TEXT, "Pendiente").click()
                        time.sleep(2)
                        texto_ingresar = "El ticket ha sido creado fuera del horario laboral. Será asignado y atendido por el personal de TI cuando vuelvan a sus labores."
                        text_area_motivo = driver.find_element(By.XPATH, "//textarea[@name='attr_pending_reason']")
                        text_area_motivo.send_keys(texto_ingresar)
                        time.sleep(3)
                        driver.find_element(By.XPATH, "//button//span[text()='Pendiente']").click()
                        time.sleep(3)

                        #buscar_incidentes()
                        #time.sleep(3)
                        break
                else:
                    #buscar_incidentes()
                    time.sleep(3)
                    break
            except StaleElementReferenceException:
                logging.error("Elemento obsoleto encontrado, reintentando...")
                time.sleep(2)
                intentos+=1
                # buscar_incidentes()
            except NoSuchElementException:
                logging.error("Elemento no encontrado, reintentando...")
                time.sleep(2)
                intentos+=1
                #buscar_incidentes()
            except Exception as e:
                logging.error(f"Error inesperado: {e}")
                break
        
        if intentos == max_intentos:
            logging.error("Máximo número de reintentos alcanzado. No se pudo completar la operación.")

    revisar_requerimientos()
    revisar_incidentes()

    logging.info("Finalizando script - cerrando driver")
    logging.info("")
    driver.quit()

# Script que se ejecuta cada 5 minutos
while True:
    hora = datetime.now()

    if (hora.weekday() >= 5 or (hora.hour == 13 and hora.minute >= 20) or hora.hour > 12 or hora.hour < 8):
        revisarTicketsFueraHorarioLaboral()

    time.sleep(300)
