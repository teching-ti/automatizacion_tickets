from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException
from selenium.webdriver.chrome.options import Options
from datetime import datetime
import time

def revisarTicketsFueraHorarioLaboral():
    chrome_options = Options()

    #las siguientes tres líneas harán que chrome se ejecute en modo headless, de esa forma no será visible

    #chrome_options.add_argument("--headless")  
    #chrome_options.add_argument("--no-sandbox")
    #chrome_options.add_argument("--disable-dev-shm-usage")

    #instancia para instalar el driver cada vez que se realice la ejecución
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    # enlace
    itop_url = ""
    username = ""
    password = ""
    
    # instancia del driver al enlace
    driver.get(itop_url)
    driver.find_element(By.ID, "user").send_keys(username)
    driver.find_element(By.ID, "pwd").send_keys(password)
    driver.find_element(By.XPATH, "//input[@type='submit' and @value='Entrar']").click()

    # funcion para dirigirse a la página de requerimientos
    def buscar_requerimientos():
        driver.find_element(By.ID, "AccordionMenu_RequestManagement").click()
        time.sleep(1)
        driver.find_element(By.LINK_TEXT, "Búsqueda de Requerimientos de Usuario").click()
        time.sleep(1)
        driver.find_element(By.CLASS_NAME, 'sfb_header').click()
        time.sleep(2)

    # funcion para dirigirse a la página de incidentes
    def buscar_incidentes():
        driver.find_element(By.ID, "AccordionMenu_IncidentManagement").click()
        time.sleep(1)
        driver.find_element(By.LINK_TEXT, "Búsqueda por Incidentes").click()
        time.sleep(1)
        driver.find_element(By.CLASS_NAME, 'sfb_header').click()
        time.sleep(2)

    # funcion para revisar los requerimientos
    def revisar_requerimientos():
        max_intentos = 5
        intentos = 0

        while intentos<max_intentos:
            print("Ejecucion req")
            # se dirige a la página donde se encuentran los requerimientos
            buscar_requerimientos()

            try:
                tabla = driver.find_element(By.CLASS_NAME, "listResults")
                # se accede a los 5 primeros registros de la tabla "seccion superior"
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
                        #print(f"Ticket {id_ticket} está fuera del horario laboral, se procede a colocar como 'Esperando Aprobación'")
                        r.find_element(By.LINK_TEXT, id_ticket).click()
                        time.sleep(3)
                        driver.find_element(By.CLASS_NAME, "actions_menu").click()
                        time.sleep(3)
                        driver.find_element(By.CSS_SELECTOR, "[data-action-id='ev_wait_for_approval']").click()
                        time.sleep(3)
                        driver.find_element(By.XPATH, "//button[@type='submit']/span[text()='Esperando Aprobación']/..").click()
                        time.sleep(3)

                        #buscar_requerimientos()
                        #time.sleep(3)
                        break
                else:
                    #buscar_requerimientos()
                    time.sleep(3)
                    break
            except StaleElementReferenceException:
                print("Elemento obsoleto encontrado, reintentando...")
                time.sleep(2)
                intentos+=1
                # buscar_requerimientos()
            except NoSuchElementException:
                print("Elemento no encontrado, reintentando...")
                time.sleep(2)
                intentos+=1
                # buscar_requerimientos()
            except Exception as e:
                print(f"Error inesperado: {e}")
                break
        print(intentos)
        if intentos == max_intentos:
            print("Máximo número de reintentos alcanzado. No se pudo completar la operación.")


    def revisar_incidentes():
        max_intentos = 5
        intentos = 0

        while intentos<max_intentos:
            print("Ejecucion inc")
            # se dirige a la página donde se encuentran los incidentes
            buscar_incidentes()

            try:
                tabla = driver.find_element(By.CLASS_NAME, "listResults")
                # se accede a los 5 primeros registros de la tabla "seccion superior"
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
                        #print(f"Ticket {id_ticket} está fuera del horario laboral, se procede a colocar como 'Esperando Aprobación'")
                        r.find_element(By.LINK_TEXT, id_ticket).click()
                        time.sleep(3)
                        driver.find_element(By.CLASS_NAME, "actions_menu").click()
                        time.sleep(3)
                        driver.find_element(By.LINK_TEXT, "Asignar").click()
                        time.sleep(2)
                        driver.find_element(By.XPATH, "//select[@name='attr_agent_id']").click()
                        time.sleep(2)
                        driver.find_element(By.XPATH, "//option[@value='113']").click()
                        time.sleep(2)
                        driver.find_element(By.XPATH, "//button//span[text()='Asignar']").click()
                        time.sleep(2)
                        driver.find_element(By.CLASS_NAME, "actions_menu").click()
                        time.sleep(2)
                        driver.find_element(By.LINK_TEXT, "Pendiente").click()
                        time.sleep(2)
                        texto_ingresar = "Ticket creado fuera del horario laboral"
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
                print("Elemento obsoleto encontrado, reintentando...")
                time.sleep(2)
                intentos+=1
                # buscar_incidentes()
            except NoSuchElementException:
                print("Elemento no encontrado, reintentando...")
                time.sleep(2)
                intentos+=1
                #buscar_incidentes()
            except Exception as e:
                print(f"Error inesperado: {e}")
                break
        print(intentos)            
        if intentos == max_intentos:
            print("Máximo número de reintentos alcanzado. No se pudo completar la operación.")

    revisar_requerimientos()
    revisar_incidentes()

    print("Finalizando script - cerrando driver")

    driver.quit()

# una vez ejecutado el script se mantendrá siempre activo
while True:
    # se obtiene la fecha y hora actual
    hora = datetime.now()

    # condicion para ejecutar las funcionalidades principales
    # solo se ejecutará sábados, domingos y cualquier día después de las 18 horas
    if(hora.weekday()>=5 or hora.hour<8 or hora.hour>=18):#18
        revisarTicketsFueraHorarioLaboral()
        
    # se ejecutará cada 5 minutos
    time.sleep(300)