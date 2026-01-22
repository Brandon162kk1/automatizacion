#-- Froms --
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from datetime import datetime, timedelta
from Sunat.validar_factura import consultarValidezSunat,login_sunat
from Birlik.cancelar_cuotas import cancelar_y_agregar_cuota, agregar_comprobante_pago,url_cuotas,url_cuotas_canceladas,url_datos_para_cancelar_cuotas
from Apis.api_birlik import consultarAPI
from GoogleChrome.chromeDriver import abrirDriver,crearCarpetas
from GoogleChrome.fecha_y_hora import get_timestamp,get_fecha_hoy
from selenium.webdriver.support.ui import Select
#-- Imports --
import subprocess
import os
import time
import pandas as pd
import sys
import random
import shutil

#----- Variables de Entorno -------
url_Positiva = os.getenv("url_Positiva")
username = os.getenv("usernamePositiva")
password = os.getenv("passwordPositiva")
# Lista de IDs de compañía
ids_compania = [12,13,14,36,38]
#----- Carpeta de la Compañia -------
nombre_carpeta_compañia = f"Positiva_{get_timestamp()}"
#-------------- Lock File -------------
# Ruta del directorio compartido entre contenedores
SYNC_DIR = "/app/sync"
# Asegurar que exista dentro del volumen (solo la primera vez)
os.makedirs(SYNC_DIR, exist_ok=True)
# Archivo de lock compartido
LOCK_FILE = os.path.join(SYNC_DIR, "session.lock")

def acquire_lock():
    try:
        # Intentar crear el archivo de lock
        fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_RDWR)
        os.write(fd, str(os.getpid()).encode())
        os.close(fd)
        return True
    except FileExistsError:
        # Ya existe => alguien más tiene el lock
        return False

def release_lock():
    try:
        os.remove(LOCK_FILE)
        print("🔓 Lock liberado.")
    except FileNotFoundError:
        pass

def wait_for_lock():
    print("🔒 Esperando que se libere el lock...")
    while True:
        if not os.path.exists(LOCK_FILE):
            if acquire_lock():
                print("✅ Lock adquirido.")
                return True
        time.sleep(5)  # espera 5 segundos antes de volver a intentar

def parse_fecha(fecha_raw):
    # Si ya es Timestamp, la retorna igual
    if isinstance(fecha_raw, pd.Timestamp):
        return fecha_raw
    fecha_str = str(fecha_raw)
    # Si la fecha tiene "-" en la posición 4, probablemente es "YYYY-MM-DD"
    if "-" in fecha_str and fecha_str[4] == "-":
        # Formato ISO (YYYY-MM-DD), NO uses dayfirst
        return pd.to_datetime(fecha_str, dayfirst=False)
    else:
        # Formato DD/MM/YYYY o similar, usa dayfirst
        return pd.to_datetime(fecha_str, dayfirst=True)

def procesar_fila(driver, wait,row,ruta_carpeta_facturas, ruta_carpeta_comprobante):
    
    tipo_doc_birlik = str(row["tipoDocumento"]).strip()
    num_poliza_birlik = str(row["numeroPoliza"]).strip()
    num_doc_birlik = str(row["numeroDocumento"]).strip()
    id_cuota_birlik = str(row["id_Cuota"]).strip()
    fk_Cliente = str(row["fk_Cliente"]).strip()
    compania_birlik =str(row["fK_Compania"]).strip()
    numero_proforma_birlik =str(row["codigoCuota"]).strip()
    importe_total_birlik = str(row["importe"]).strip()
    estadoCuota_birlik = str(row["estadoCuota"]).strip()
    ramo_birlik = int(str(row["fk_Ramo"]).strip())
    vigencia_inicio_raw = row["vigenciaInicio"] #String
    vigencia_fin_raw = row["vigenciaFin"] #String

    #-- Convertimos la fechas a formato datetime ---
    fecha_inicio = datetime.strptime(vigencia_inicio_raw, "%d/%m/%Y")
    fecha_fin = datetime.strptime(vigencia_fin_raw, "%d/%m/%Y")

    resultado_importe = False
    resultado_sunat = False
    resultado_birlik = False
    resultado_estado = None
    resultado_accion = ""

    # if estadoCuota_birlik == "Pendiente":
    #      return f"Pagina Web en Mantenimiento" if resultado_importe else "Pagina Web en Mantenimiento" ,"Pagina Web en Mantenimiento" if resultado_sunat else "Pagina Web en Mantenimiento" ,"Pagina Web en Mantenimiento" if resultado_birlik else "Pagina Web en Mantenimiento" , "" if resultado_estado else "Pagina Web en Mantenimiento", "" if resultado_accion else "Pagina Web en Mantenimiento"

    # if estadoCuota_birlik not in ("11138","10342","11325","10783","8013","5925","564","554","552","551","11308","11324","8114","7884","7882","5924"):
         # return f"Pagina Web en Mantenimiento" if resultado_importe else "Pagina Web en Mantenimiento" ,"Pagina Web en Mantenimiento" if resultado_sunat else "Pagina Web en Mantenimiento" ,"Pagina Web en Mantenimiento" if resultado_birlik else "Pagina Web en Mantenimiento" , "" if resultado_estado else "Pagina Web en Mantenimiento", "" if resultado_accion else "Pagina Web en Mantenimiento"

    #--- 1era opcion---
    fecha_menos_7dias = fecha_inicio - timedelta(days=7) # Restar 7 días (1 semana) o una_semana_antes
    # Sumar 350 días a la fecha de 1 semana antes o fecha_350_dias_despues
    fecha_350_dias_despues = fecha_menos_7dias + timedelta(days=359)
    # Convertir a string para mostrar o guardar
    fecha_opcion_1_inicio = fecha_menos_7dias.strftime("%d/%m/%Y")
    fecha_opcion_1_final = fecha_350_dias_despues.strftime("%d/%m/%Y")

    #---2da opcion---
    dia_fecha_hoy = get_fecha_hoy().date() # --> 2025-08-04
    # Restar 359 días
    fecha_hoy_menos_359 = dia_fecha_hoy - timedelta(days=359)
    # Convertir ambas a string con formato dd/mm/yyyy
    fecha_opcion_2_final = dia_fecha_hoy.strftime("%d/%m/%Y")
    fecha_opcion_2_inicio = fecha_hoy_menos_359.strftime("%d/%m/%Y")

    #-- 3 era opcion ---
    fecha_opcion_3_inicio = fecha_inicio.strftime("%d/%m/%Y")
    fecha_opcion_3_final = fecha_fin.strftime("%d/%m/%Y")

    #-- 4ta opcion ---
    fecha_1mesantes_vigIni = fecha_inicio - timedelta(days=30)
    fecha_opcion_4_inicio = fecha_1mesantes_vigIni.strftime("%d/%m/%Y")
    fecha_opcion_4_fi = fecha_1mesantes_vigIni + timedelta(days=359)
    fecha_opcion_4_final = fecha_opcion_4_fi.strftime("%d/%m/%Y")

    # Al inicio del script
    fechas_inicio = [fecha_opcion_1_inicio,fecha_opcion_2_inicio,fecha_opcion_3_inicio,fecha_opcion_4_inicio]
    fechas_fin = [fecha_opcion_1_final,fecha_opcion_2_final,fecha_opcion_3_final,fecha_opcion_4_final]

    print(f"Cliente con {tipo_doc_birlik} # {num_doc_birlik} ")


    if ramo_birlik == 55 or ramo_birlik == 56:
        ruc_emisor = '20454073143'
        nom_compania = "La Positiva Vida Seguros y Reaseguros (Pension o Vida Ley)"
    elif ramo_birlik == 54:
        ruc_emisor = '20601978572'
        nom_compania = "LA POSITIVA EPS (Salud)"
    elif ramo_birlik not in [54, 55, 56]:
        ruc_emisor = '20100210909'
        nom_compania = "La Positiva Seguros y Reaseguros (Otros)"
    else:
        print(f"Valor inesperado en CompaniasMismaTrazabilidad: '{compania_birlik}'. Se aborta el procesamiento de esta fila.")
        return
  
    print(f"{nom_compania} con RUC: {ruc_emisor}")

    try:

        ventana_original = driver.current_window_handle
        
        intentos = 0
        max_intentos = len(fechas_inicio)

        while intentos < max_intentos:
            intentos += 1

            try:

                if intentos > 1:
                    print("🔁 Recargando página")
                    driver.refresh()
                    time.sleep(5)

                print(f"🔄 Intento número {intentos}")

                fecha_ini_actual = fechas_inicio[intentos - 1]
                fecha_fin_actual = fechas_fin[intentos - 1]

                print(f"📅 Fecha Desde: {fecha_ini_actual} Hasta: {fecha_fin_actual}")

                combo_tipo_doc = wait.until(EC.element_to_be_clickable((By.NAME, "lscTipoDoc")))
                combo_tipo_doc.click()
        
                if tipo_doc_birlik == "RUC":
                    valor= "2"
                else:
                    valor= "1"
                
                select_tipo_doc = Select(combo_tipo_doc)
                select_tipo_doc.select_by_value(valor)
        
                ruc_field = wait.until(EC.presence_of_element_located((By.NAME, "tctNumDoc")))
                ruc_field.clear()
                ruc_field.send_keys(num_doc_birlik)
       
                fecha_input = wait.until(EC.presence_of_element_located((By.NAME, "sFechai")))        
                fecha_input.clear()       
                fecha_input.send_keys(fecha_ini_actual)

                fecha_input_fin = wait.until(EC.presence_of_element_located((By.NAME, "sFechaf")))        
                fecha_input_fin.clear()      
                fecha_input_fin.send_keys(fecha_fin_actual)

                driver.execute_script("Buscar();")

                time.sleep(3)
    
                tablas_ids = ["tabla1", "tabla2", "tabla3"]
        
                proforma_encontrada = False         
                estado_valor = None                 
                comprobante_valor = None            
                importe_valor = 0                   
              
                for tabla_id in tablas_ids:

                    try:
                
                        print(f"⌛ Iniciando la búsqueda en: {tabla_id}")   

                        wait.until(EC.visibility_of_element_located((By.ID, tabla_id)))

                        rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, f"//*[@id='{tabla_id}']//tr[contains(@class,'row11')]")))
                        print(f"✅ {tabla_id} cargada con {len(rows)} filas.")

                        for row in rows:

                            cells = row.find_elements(By.TAG_NAME, "td")
            
                            if len(cells) >= 9:

                                numero = cells[8].text.strip()
                
                                if numero == numero_proforma_birlik:

                                    proforma_encontrada = True

                                    cta_cobrar = cells[12].text.strip()
                                    estado = cells[14].text.strip()
                                    fecha_facturacion = cells[6].text.strip()   
                                    fecha_estado = cells[15].text.strip()           
                                    nro_serie = cells[17].text.strip()
                                    nro_comprobante = cells[18].text.strip()

                                    estado_valor = estado
                                    resultado_estado = estado
                                    comprobante_valor = str(nro_serie)+ "-" + str(nro_comprobante)

                                    importe_valor = float(cta_cobrar.replace(',', '')) 

                                    print(f"Importe de Birlik: {float(importe_total_birlik)} -- Importe de la Compañía : {float(importe_valor)}")

                                    diferencia = abs(float(importe_valor) - float(importe_total_birlik))
                                    if diferencia > 0.05:
                                        print("❌ Los importes No coinciden")
                                    else:
                                        resultado_importe = True
                                        print("✅ Los importes Coinciden")

                                    if estado_valor == "Pendiente":
                                        print(f"⚠️ Estado de la Cuota {numero_proforma_birlik} como '{estado_valor}'.")
                                        resultado_estado = estado_valor
                                        resultado_accion = f'Sin Observación'
                                        break

                                    if "ANUL" in estado_valor.upper():
                                        print(f"⚠️ Estado de la Cuota {numero_proforma_birlik}: '{estado_valor}'. No se procesa esta cuota.")
                                        resultado_estado = estado_valor
                                        resultado_accion = f'=HYPERLINK("{url_cuotas}{fk_Cliente}", "Anular Cuota")'
                                        break

                                    print(f"✅ Cuota encontrada con estado: '{estado_valor}'")

                                    if len(cells) > 20:

                                        accion_cell = cells[20]
                        
                                        try:

                                            ventanas_antes = len(driver.window_handles)
                                            ventana_principal_positiva = driver.current_window_handle

                                            accion_img = accion_cell.find_element(By.TAG_NAME, "img")
                                            accion_img.click()
                                            print(f"🖱️ Clic en la acción para la cuota: {numero_proforma_birlik}")
                            
                                            wait.until(lambda d: len(d.window_handles) > ventanas_antes)
                            
                                            driver.switch_to.window(driver.window_handles[-1])
                                            print("🔄 Cambiado a la nueva ventana/pestaña.")

                                            wait.until(lambda d: len(d.window_handles) > 1)
                                            nuevas_ventanas = [w for w in driver.window_handles if w != ventana_original]
                                            driver.switch_to.window(nuevas_ventanas[0])

                                            destino_factura = f"{ruta_carpeta_facturas}/{num_poliza_birlik}_{comprobante_valor}"
                                            ruta_factura = os.path.join(ruta_carpeta_facturas, f"{num_poliza_birlik}_{comprobante_valor}.pdf")

                                            #------------------------------
                                            time.sleep(3)
                                            subprocess.run(["xdotool", "search", "--name", "Save File", "windowactivate", "windowfocus"])
                                            print("💡 Se hizo FOCO en la nueva ventana de dialogo de Linux Debian")
                                            time.sleep(3)                              
                                            subprocess.run(['xdotool', 'type',"--delay", "100", destino_factura])
                                            print("📄 Se escribió el nombre del archivo")
                                            time.sleep(3)
                                            subprocess.run(['xdotool', 'key', 'Return'])
                                            print("🖱️ Se dio Enter para confirmar")
                                            time.sleep(3)
                                            #------------------------------

                                            try:
                                                driver.switch_to.window(ventana_original)
                                                print("🔄 Volvimos a la ventana original")
                                            except Exception as e:
                                                print(f"❌ No se pudo volver a la ventana original, Motivo: {e}")
                            
                                            time.sleep(3)

                                            fechas = [fecha_estado,fecha_facturacion]

                                            for fecha in fechas:

                                                nombre_imagen_sunat = f"{numero_proforma_birlik}_{num_poliza_birlik}.png"
                                                ruta_imagen_sunat = os.path.join(ruta_carpeta_comprobante, nombre_imagen_sunat)
                                                resultado = consultarValidezSunat(driver,wait,ruc_emisor,tipo_doc_birlik,num_doc_birlik,comprobante_valor,fecha,cta_cobrar,ruta_imagen_sunat)

                                                driver.switch_to.window(ventana_principal_positiva)
                                                print("🔄 Volviendo a la ventana de la CIA")

                                                if resultado is None:
                                                    resultado_accion = f'=HYPERLINK("{login_sunat}", "Sunat Bloqueado")'
                                                    break
                                                elif resultado:

                                                    resultado_sunat = True

                                                    if estadoCuota_birlik == "Pendiente-comprobante":
                                                        print("📤 Subiendo comprobante a Birlik...")
                                                        agregar_comprobante_pago(driver,wait,id_cuota_birlik,ruta_factura)
                                                        resultado_accion = "Factura Enviada Anteriormente"
                                                    else:
                                                        print("📤 Subiendo todos los documentos a Birlik...")
                                                        cancelar_y_agregar_cuota(driver,wait,id_cuota_birlik,comprobante_valor,fecha,ruta_factura,ruta_imagen_sunat,resultado_importe)
                                                        resultado_accion = f'=HYPERLINK("{url_cuotas_canceladas}{fk_Cliente}", "Enviar Factura")'
         
                                                    resultado_birlik = True

                                                    break

                                                else:
                                                    resultado_accion = f'=HYPERLINK("{login_sunat}", "Ver Sunat")'
                                                    continue
                                            
                                            break

                                        except Exception as e:
                                            print(f"No se pudo hacer clic en la imagen de acción,Flujo Terminado con estado {estado_valor}")

                        if proforma_encontrada:
                            break

                    except Exception as e:
                        pass

                if proforma_encontrada:
                    break
                else:
                    print("❌ No se encontró la Proforma")

            except Exception as e:
                print(f"❌ Error durante el intento {intentos}")

    except Exception as e:
        print(f"❌ Error Procesando toda la fila, Revisar: {e}")
    finally:
        driver.refresh()
        return f"Coinciden" if resultado_importe else f"No Existe en la CIA" ,"Válido" if resultado_sunat else f"No Existe en la CIA" ,"Cuota Cancelada" if resultado_birlik else "Cuota Pendiente" , f"Cancelado" if resultado_estado else f"No Existe en la CIA" , resultado_accion
    
def validar_pagina(driver):

    asunto = ""

    try:

        # Validar si aparece el mensaje de error en el body
        if "The requested URL was rejected. Please consult with your administrator." in driver.page_source:
            asunto = "❌ Página web de La Positiva fuera de Servicio"
            return False,asunto

        # Validar si aparece el campo de usuario
        user_field = WebDriverWait(driver,5).until(
            EC.presence_of_element_located((By.NAME, "txtUsuario"))
        )
        if user_field:
            asunto = "❌ Página web de La Positiva fuera de Servicio"
            return False,asunto

    except TimeoutException:
        print("✅ Continuamos")
        return True,asunto

def escribir_lento(elemento, texto,min_delay=0.7, max_delay=0.9):
    """Envía texto carácter por carácter con retrasos aleatorios."""
    for letra in texto:
        elemento.send_keys(letra)
        time.sleep(random.uniform(min_delay, max_delay))

def mover_y_hacer_click_simple(driver, elemento, steps=6, pause_between=0.06):
    """
    Mueve el mouse en 'steps' pasos hacia el centro del elemento y hace click.
    driver: tu instancia de webdriver
    elemento: WebElement destino
    """
    action = ActionChains(driver)
    # asegurarnos que el elemento esté visible en pantalla
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
    time.sleep(random.uniform(0.15, 0.45))

    # posiciona el mouse sobre el elemento (move_to_element genera mouseover)
    action.move_to_element(elemento).pause(random.uniform(0.05, 0.18)).perform()

    # pequeños movimientos aleatorios alrededor antes del click
    for _ in range(steps):
        offset_x = random.randint(-6, 6)
        offset_y = random.randint(-6, 6)
        action.move_by_offset(offset_x, offset_y).pause(pause_between)
    # volver al elemento y click
    action.move_to_element(elemento).pause(random.uniform(0.08, 0.2)).click().perform()

def main():
      
    display_num = os.getenv("DISPLAY_NUM", "0")
    os.environ["DISPLAY"] = f":{display_num}"

    driver, wait = abrirDriver()

    driver.get(url_Positiva)
    print("✅ Ingresando a la URL")
    login_exitoso = False

    for intento in range(2):
        print(f"🔄 Intento de login número {intento + 1}...")

        try:

            time.sleep(3)
            user_field = wait.until(EC.presence_of_element_located((By.ID, "b5-Input_User")))
            user_field.clear()

            mover_y_hacer_click_simple(driver, user_field)
            time.sleep(random.uniform(0.97, 0.99))

            escribir_lento(user_field, username, min_delay=0.97, max_delay=0.99)
            print("✅ Digitando el Username")

            time.sleep(1 + random.random() * 1.5)

            password_field = wait.until(EC.presence_of_element_located((By.ID, "b5-Input_PassWord")))
            password_field.clear()

            mover_y_hacer_click_simple(driver, password_field)
            time.sleep(random.uniform(0.97, 0.99))

            escribir_lento(password_field, password, min_delay=0.97, max_delay=0.99)
            print(f"✅ Digitando el Password '{password}'.")

            time.sleep(1 + random.random() * 1.5)

            login_button = wait.until(EC.element_to_be_clickable((By.ID, "b5-btnAction")))
            mover_y_hacer_click_simple(driver, login_button)
            print("🖱️ Clic en Inicar Sesión")

            #time.sleep(5)

            try:
                popup_text = WebDriverWait(driver,7).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(),'Usuario o contraseña incorrectos')]")))
                if popup_text:
                    print("❌ Usuario o contraseña incorrectos")
                    aceptar_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//button[.//span[text()='Aceptar']]")))
                    aceptar_btn.click()
                    time.sleep(2)

                    driver.refresh()
                    time.sleep(2)
                    continue

            except:
                pass

            login_exitoso = True
            break

        except Exception as e:
            print(f"❌ Error durante el intento {intento+1}: {e}")
            driver.refresh()
            time.sleep(2)

    time.sleep(3)

    if not login_exitoso:
        print("❌ No se pudo iniciar sesión después de 2 intentos.")
        print("⌛ Esperando 30 minutos para poder ingresar de nuevo.")
        #release_lock()
        time.sleep(1800)
        sys.exit()

    wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'menu-item')]//span[normalize-space()='Autogestión']/parent::div")))
    autogestion = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class,'menu-item')]//span[normalize-space()='Autogestión']/parent::div")))

    driver.execute_script("arguments[0].click();", autogestion)
    print("✅ Login exitoso")
    print("🖱️ Clic en Autogestión")

    ov = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='OV']")))
    ov.click()
    print("🖱️ Clic en 'OV' ")

    wait.until(lambda d: len(d.window_handles) > 1)

    driver.switch_to.window(driver.window_handles[-1])

    try:
        alert = WebDriverWait(driver,5).until(EC.alert_is_present())
        print(f"⚠️ Alerta presente: {alert.text}")
        alert.accept()
        print("✅ Alerta aceptada")
    except:
        print("✅ No apareció ninguna alerta")
                
    resultado,asunto = validar_pagina(driver)    

    if not resultado:

        print(asunto)
        print("⌛ Esperando 30 minutos para poder ingresar de nuevo.")
        #release_lock()
        time.sleep(1800) #---> Espera 30 min de forma manual
        sys.exit()

    estados_de_cuenta_img = wait.until(EC.presence_of_element_located((By.ID, f"stUIUserOV31_img")))
    print("🖱️ Mouse sobre la imagen 'Estado de Cuenta'.")

    action = ActionChains(driver)
    action.move_to_element(estados_de_cuenta_img).perform()

    span_element = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'stsp') and contains(text(), 'Estado') and contains(text(), 'cuenta')]")))

    id_dinamico = span_element.get_attribute("id")

    estado_de_cuenta_link = wait.until(EC.element_to_be_clickable((By.ID, id_dinamico)))
    estado_de_cuenta_link.click()
    print("🖱️ Clic en 'Estado de Cuenta'.")

    time.sleep(3)

    while True:

        try:
            #release_lock()

            todos_los_datos = consultarAPI(url_datos_para_cancelar_cuotas,ids_compania)

            if not todos_los_datos:
                raise Exception("No se recibió información de ninguna compañía.")

            log_path,ruta_salida_API,ruta_salida,ruta_maestro,nombre_hoja, ruta_carpeta_facturas, ruta_carpeta_comprobante, ruta_carpeta_errores,carpeta_compañia,carpeta_principal = crearCarpetas(todos_los_datos,nombre_carpeta_compañia,tipo=2,cia_a_verificar=None)

            print("\n📁 Iniciando procesamiento del Excel..")           

            try:
                df = pd.read_excel(ruta_salida_API, engine="openpyxl", dtype={"numeroDocumento": str})
            except Exception as e:
                raise Exception(f"Error al leer el archivo Excel: {e}")

            df["Importe"] = ""
            df["Sunat"] = ""
            df["Birlik"] = ""
            df["Estado"] = ""
            df["Acción"] = ""
            
            total_filas = len(df)

            for index, row in df.iterrows():
                print(f"\n--- Procesando fila {index + 2} de {total_filas + 1} ---")
                #wait_for_lock()
                try:
                    importe_estado,sunat_estado,birlik_estado,estado_estado,accion_estado = procesar_fila(driver,wait,
                        row,ruta_carpeta_facturas, ruta_carpeta_comprobante)

                    df.at[index, "Importe"] = importe_estado
                    df.at[index, "Sunat"] = sunat_estado
                    df.at[index, "Birlik"] = birlik_estado
                    df.at[index, "Estado"] = estado_estado
                    df.at[index, "Acción"] = accion_estado

                    df.to_excel(ruta_salida, index=False)

                except Exception as e:
                    print(f"❌ Error en fila {index}, {e}")
                    df.to_excel(ruta_salida, index=False)
                finally:
                    print(f"✅ Fila {index} guardada correctamente.")

                time.sleep(3)

            #df.to_excel(ruta_salida, index=False)     
            #guardar_excel_con_formato(ruta_salida,'Sheet1')

        except Exception as e:
            print(f"\n🟡 Proceso Detenido, por :{e}")
        finally:
            #release_lock()
            driver.quit()
            if os.path.exists(carpeta_principal):
                shutil.rmtree(carpeta_compañia)
                print("🧹 Carpeta eliminada correctamente")
            else:
                print("⚠️ La carpeta no existe")
            print(f"\n✅ Flujo finalizado, Intentando de nuevo en 10 segundos.")
            time.sleep(10)

if __name__ == "__main__":
    main()