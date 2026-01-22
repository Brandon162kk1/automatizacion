#-- Imports ---
import re
import sys
import subprocess
import time
import os
import pandas as pd
import shutil
#import pyautogui  # Solo si necesitas automatizar ventanas nativas
#-- Froms ----
from selenium.webdriver.common.by import By
from selenium.webdriver.support.expected_conditions import invisibility_of_element_located
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import Select
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from cuotas_Sanitas_Crecer import buscaryRegistrarenCrecer
from Sunat.validar_factura import consultarValidezSunat,login_sunat
from Birlik.cancelar_cuotas import agregar_comprobante_pago,cancelar_y_agregar_cuota,url_cuotas,url_cuotas_canceladas,url_datos_para_cancelar_cuotas
from Apis.api_birlik import consultarAPI
from GoogleChrome.chromeDriver import abrirDriver, crearCarpetas
from GoogleChrome.fecha_y_hora import get_timestamp,get_fecha_hoy

#--------- COMPAÑÍA SANITAS ---------------
ruc_sanitas = '20523470761'
#--------- COMPAÑÍA SANITAS PROTECTA ------
#login_url_sanitas_protecta = 'https://protectasecurity.sanitasperu.com/'
login_url_sanitas_protecta = os.getenv("login_url_sanitas_protecta")
ruc_sanitas_protecta = '20517207331'
#--------- COMPAÑÍA SANITAS CRECER --------
#login_url_sanitas_crecer = 'https://crecerseguros.sanitasperu.com/'
login_url_sanitas_crecer = os.getenv("login_url_sanitas_crecer")
ruc_sanitas_crecer = '20600098633'
# Lista de IDs de compañía
ids_compania = [5,29,31]
#----- Variables de Entorno -------
username_sanitas = os.getenv("usernameSanitas")
password_sanitas = os.getenv("passwordSanitas")
#----- Carpeta de la Compañia -------
nombre_carpeta_compañia = f"Sanitas_SCTR_{get_timestamp()}"

def click_descarga_opcion(driver, destino_factura, boton_descarga,numero_poliza,ruta_carpeta_errores):
    
    try:

        driver.execute_script("arguments[0].click();", boton_descarga)
        print("✅ Se hizo clic con JS en el botón de descarga.")

        print("⌛ Esperando la ventana descarga de Linux Debian...")
        time.sleep(2)

        subprocess.run(["xdotool", "search", "--name", "Save File", "windowactivate", "windowfocus"])
        print("💡 Se hizo FOCO en la nueva ventana de dialogo de Linux Debian")
        time.sleep(2)

        subprocess.run(["xdotool", "type","--delay", "100", destino_factura])
        print("📄 Se escribió el nombre del archivo")

        time.sleep(2)

        subprocess.run(["xdotool", "key", "Return"])
        print("🖱️ Se presionó Enter para confirmar la descarga.")

        time.sleep(2)
        return True

    except Exception as ex:
        print("❌ Error durante el flujo de descarga:", ex)
        driver.save_screenshot(f"{ruta_carpeta_errores}/{numero_poliza}_ventanalinux.png")
        return False

def procesar_fila(row,ruta_carpeta_facturas, ruta_carpeta_comprobante, ruta_carpeta_errores,carpeta_compañia):

    #--Extraer valores y quitar espacios en blanco
    numero_poliza = str(row["numeroPoliza"]).strip()
    tipo_doc_birlik = str(row["tipoDocumento"]).strip()
    numero_ruc = str(row["numeroDocumento"]).strip()
    id_cuota = str(row["id_Cuota"]).strip()
    fk_Cliente = str(row["fk_Cliente"]).strip() 
    fk_compania = str(row["fK_Compania"]).strip()
    numero_proforma = str(row["codigoCuota"]).strip()
    importe_total_birlik = str(row["importe"]).strip()
    estadoCuota_birlik = str(row["estadoCuota"]).strip()
    #--------------

    resultado_importe = False
    resultado_sunat = False
    resultado_birlik = False
    resultado_ocr = False
    resultado_estado = None
    resultado_accion = ""

    # if id_cuota not in ("7833"):
    #     return f"Pagina Web en Mantenimiento" if resultado_importe else "Pagina Web en Mantenimiento" ,"Pagina Web en Mantenimiento" if resultado_sunat else "Pagina Web en Mantenimiento" ,"Pagina Web en Mantenimiento" if resultado_birlik else "Pagina Web en Mantenimiento" ,"Pagina Web en Mantenimiento" if resultado_ocr else "Pagina Web en Mantenimiento", "" if resultado_estado else "Pagina Web en Mantenimiento", "" if resultado_accion else "Pagina Web en Mantenimiento"

    if fk_compania == '29':
        #--- SANITAS
        nombre_Compania = "Sanitas"
        url_general = login_url_sanitas_protecta
        ruc_compania = ruc_sanitas
    elif fk_compania == '31':
        #--- SANITAS PROTECTA
        nombre_Compania = "Sanitas Protecta"
        url_general = login_url_sanitas_protecta
        ruc_compania = ruc_sanitas_protecta
    else:
        #--- SANITAS CRECER
        nombre_Compania = "Sanitas Crecer"
        url_general = login_url_sanitas_crecer
        ruc_compania = ruc_sanitas_crecer

    print(f"Compañía {nombre_Compania} - RUC: {ruc_compania}")

    display_num = os.getenv("DISPLAY_NUM", "0")  # fallback = 0
    os.environ["DISPLAY"] = f":{display_num}"

    driver, wait = abrirDriver()
    
    try:
        # ------------------ Inicio del Flujo de Automatización ------------------
        driver.get(url_general) 
        print("✅ Ingresando a la URL")
        
        # 2. Rellenar el formulario de login en Sanitas
        user_input = wait.until(EC.presence_of_element_located((By.ID, "Login")))
        user_input.clear()
        user_input.send_keys(username_sanitas)
        print("✅ Digitando el Username")
        
        pass_input = wait.until(EC.presence_of_element_located((By.ID, "Password")))
        pass_input.clear()
        pass_input.send_keys(password_sanitas)
        print("✅ Digitando el Password")
        
        ingresar_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(text(),'Ingresar')]")
        ))
        ingresar_btn.click()
        print("🖱️ Se hizo clic en 'Ingresar' en Sanitas.")
        
        # 3. Esperar a que cargue la página Quotation/Index
        wait.until(EC.url_contains("Quotation/Index"))
        
        # 4. Clic en "Autogestión"
        autogestion_link = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(@class,'dropdown-toggle') and contains(text(),'Autogestión')]")
        ))
        autogestion_link.click()
        print("🖱️ Se hizo clic en 'Autogestión'.")
        
        # 5. Clic en "Estado de cuenta SCTR"
        estado_cuenta_link = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(text(),'Estado de cuenta SCTR')]")
        ))
        estado_cuenta_link.click()
        print("🖱️ Se hizo clic en 'Estado de cuenta SCTR'.")
        
        # # 6. Seleccionar en el dropdown el tipo de documento: RUC (value="2")
        # select_identity = wait.until(EC.presence_of_element_located((By.ID, "IdentityTypeId")))
        # select_obj = Select(select_identity)
        # select_obj.select_by_value("2")
        # print("Se seleccionó 'RUC' en el dropdown.")

        # 7. Ingresar el número de RUC
        ruc_input = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//input[@type='text' and contains(@class, 'AccountStatusList')]")
        ))
        ruc_input.clear()
        ruc_input.send_keys(numero_ruc)
        print(f"✅ Se ingresó el RUC: {numero_ruc}")
        
        try:
            # 8. Clic en el botón que despliega las opciones de estado (dropdown de status)
            status_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//div[contains(@class,'btn-group') and contains(@class,'AccountStatusList')]//button")
            ))
            status_btn.click()
            print("🖱️ Se hizo clic en el dropdown de estado.")
        except Exception as e:
            print(f"❌ Error al hacer clic en el botón {status_btn}", e)

        try:
            # 9. Seleccionar la opción "Abonada" en el dropdown
            abonada_option = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//li[.//span[contains(text(),'Abonada')]]")
            ))
            abonada_option.click()
            print("🖱️ Se seleccionó la opción 'Abonada'.")
        except Exception as e:
            print(f"❌ Error al hacer clic en el botón {abonada_option}", e)

        
        try:
            # 9.1. Seleccionar la opción "Anulada" en el dropdown
            anulada_option = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//li[.//span[contains(text(),'Anulada')]]")
            ))
            anulada_option.click()
            print("🖱️ Se seleccionó la opción 'Anulada'.")
        except Exception as e:
            print(f"❌ Error al hacer clic en el botón {anulada_option}", e)
            
        try:
            # 9.2. Seleccionar la opción "Migración" en el dropdown
            migracion_option = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//li[.//span[contains(text(),'Migración')]]")
            ))
            migracion_option.click()
            print("🖱️ Se seleccionó la opción 'Migración'.")
        except Exception as e:
            print(f"❌ Error al hacer clic en el botón {migracion_option}", e)
        
        # 9.3. Hacer clic en el body para cerrar el dropdown
        driver.find_element(By.TAG_NAME, "body").click()
        print("🖱️ Se cerró el dropdown (clic en body).")
        
        # 10. Clic en el botón "Filtrar"
        filter_btn = wait.until(EC.element_to_be_clickable((By.ID, "btnSearch")))
        filter_btn.click()
        print("🖱️ Se hizo clic en 'Filtrar'.")

        # Esperar que desaparezca el mensaje "Un momento por favor..."
        wait.until(invisibility_of_element_located((By.ID, "AccountStatusListTable_processing")))
        print("⌛ Cargando...")

        # 11. Seleccionar la opción "1000" en el dropdown de cantidad de registros
        select_elem = wait.until(EC.presence_of_element_located((By.NAME, "AccountStatusListTable_length")))
        Select(select_elem).select_by_value("1000")
        print("🖱️ Se seleccionó '1000' registros para mostrar más filas.")

        # Esperar que desaparezca el mensaje "Un momento por favor..."
        wait.until(invisibility_of_element_located((By.ID, "AccountStatusListTable_processing")))

        try:
            wait.until(
                lambda d: len(
                    d.find_elements(By.XPATH, "//table[@id='AccountStatusListTable']//tr")
                ) > 1
            )
            print("✅ La tabla tiene al menos 1 fila.")
        except TimeoutException:
            print("⏰ No se cargaron las filas en el tiempo esperado.")

        # 2. Obtener la tabla
        table = driver.find_element(By.ID, "AccountStatusListTable")

        # 3. Obtener todas las filas
        rows = table.find_elements(By.TAG_NAME, "tr")
        
        fila_encontrada = False
        comprobante_valor = None    #F002-01849092
        estado_valor = None         # Abonado, Emitida o Impresa
        documento_valor = None      # PF-SCTR-002737767 -> Contiene el Numero de Cuota
        fecha_emision_valor = None
        importe_valor = 0
   
        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) == 11:                        # Asegurarse que es una fila con 11 celdas
                fecha_emision = cells[3].text.strip()   # Fecha Emisión
                documento = cells[4].text.strip()       # PF-SCTR-002737767 -> Contiene el Numero de Cuota
                comprobante = cells[6].text.strip()     #Factura F002-01849092
                estado = cells[7].text.strip()
                importe = cells[8].text.strip()
                
                # Extrae solo números del documento
                solo_numeros_doc = re.sub(r'\D', '', documento)

                if numero_proforma in solo_numeros_doc:
                    fila_encontrada = True
                    estado_valor = estado
                    resultado_estado = estado_valor     #Le asignamos el valor del estado al Campo Estado
                    comprobante_valor = comprobante
                    documento_valor = documento
                    fecha_emision_valor = fecha_emision
                    importe_valor = importe

                    print(f"✅ Fila encontrada: Documento='{documento}', Estado='{estado}', Comprobante='{comprobante}', Importe ='{importe}'")
                    
                    if estado.lower() == "abonada":
                        print(f"✅ La cuota con proforma {numero_proforma} está abonada. Comprobante: {comprobante_valor}")
                    elif estado.lower() == 'anulada':
                        print(f"❌ La cuota con proforma {numero_proforma} está anulada. Comprobante: {comprobante_valor}")
                        resultado_accion = f'=HYPERLINK("{url_cuotas}{fk_Cliente}", "Anular Cuota")'
                    else:
                        print(f"⚠️ La cuota con proforma {numero_proforma} está pendiente (Estado='{estado}')")
                        resultado_accion = f'Sin Observación'
                    
                    break
                else:
                    resultado_importe = "Codigo Cuota Incorrecto"
        
        print(f"Importe de Birlik: {float(importe_total_birlik)} -- Importe de la Compañía : {float(importe_valor)}")

        # Validación de importes con tolerancia
        diferencia = abs(float(importe_valor) - float(importe_total_birlik))

        if diferencia > 0.05:
            print("❌ Los importes No coinciden")
        else:
            resultado_importe = True
            print("✅ Los importes Coinciden")

        # # Validación de importes
        # if float(importe_valor) == float(importe_total_birlik):
        #     resultado_importe = True
        #     print("✅ Los importes de la Compañía y Birlik son iguales")

        if not fila_encontrada:
            print(f"❌ No se encontró ninguna fila con Documento conteniendo '{numero_proforma}'")
        
        # 13. Si la cuota está abonada, navegar a "Consulta de Comprobantes de pago"
        if fila_encontrada and estado_valor.lower() == "abonada":
            print(f"✅ La cuota con proforma {numero_proforma} está abonada. Comprobante: {comprobante_valor}")
            # Si además quieres guardar el valor de Documento y Fecha Emisión, ya lo tienes almacenado
            print(f"Documento: {documento_valor} - Fecha Emisión: {fecha_emision_valor}")

            # 13.1. Clic en el menú "Autogestión"
            autogestion_link = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//a[contains(@class,'dropdown-toggle') and contains(text(),'Autogestión')]")
            ))
            autogestion_link.click()
            time.sleep(2)
            
            # 13.2. Clic en "Consulta de Comprobantes de pago"
            consulta_comprobantes_link = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//a[@href='/PaymentVouchers/LegalDocumentsSelfManagementIndex' and contains(text(),'Comprobantes')]")
            ))
            consulta_comprobantes_link.click()
            print("🌐 Navegando a 'Consulta de Comprobantes de pago'.")
            time.sleep(3)

            # Localizar el input por ID "ContractNumber"
            contract_input = wait.until(EC.presence_of_element_located((By.ID, "ContractNumber")))
            contract_input.clear()  # Limpia el campo, si es necesario
            contract_input.send_keys(numero_poliza)
            print(f"✅ Se ingresó el número de póliza: {numero_poliza}")

            # Ingresando las fechas segun la fecha de emision

            # Suponiendo que fecha_emision_valor es un string, por ejemplo:
            fecha = datetime.strptime(fecha_emision_valor, "%d/%m/%Y")
            fecha_inicio = fecha - relativedelta(months=1)
            fecha_fin = fecha + relativedelta(months=1)

            # Formatear el rango de fechas en el formato requerido: "dd/mm/yyyy - dd/mm/yyyy"
            rango_fechas = f"{fecha_inicio.strftime('%d/%m/%Y')} - {fecha_fin.strftime('%d/%m/%Y')}"
            print("📅 Rango de fechas:", rango_fechas)

            # Localizar el input de fecha y enviar el rango
            input_rango = wait.until(EC.presence_of_element_located((By.ID, "DatesRangeCreation")))
            input_rango.clear()
            input_rango.send_keys(rango_fechas)

            # Luego, hacer clic en el botón Filtrar
            filter_btn = wait.until(EC.element_to_be_clickable((By.ID, "btnSearch")))
            filter_btn.click()

            # Seleccionar "1000" registros para que se muestren todas las filas:
            select_elem = wait.until(EC.presence_of_element_located((By.NAME, "LegalDocumentSelfManagementListTable_length")))
            Select(select_elem).select_by_value("1000")
            print("🖱️ Seleccionado '1000' registros para ver más filas.")
            time.sleep(2)
    
            # Localizar la tabla:
            tabla = wait.until(EC.presence_of_element_located((By.ID, "LegalDocumentSelfManagementListTable")))
            filas_tabla = tabla.find_elements(By.TAG_NAME, "tr")
    
            fila_encontrada_descarga = False
            for fila in filas_tabla:
                celdas = fila.find_elements(By.TAG_NAME, "td")
                # Supongamos que la columna "Documento" es la 3ª (índice 2) en esta tabla.
                # Ajusta según el orden real en tu tabla.
                if len(celdas) >= 4:
                    doc_valor = celdas[2].text.strip()  # Ajusta índice si es necesario
                    if comprobante_valor in doc_valor:
                        fila_encontrada_descarga = True
                        print(f"✅ Fila encontrada en 'Consulta de Comprobantes': Documento = {doc_valor}")
                
                        # Ahora, en esa fila, queremos hacer clic en el icono que despliega el menú.
                        # Asumiremos que el icono está en la última celda.
                        celda_accion = celdas[-1]
                        
                        # 1. Localiza el ícono desplegable dentro de la celda
                        # Hacer clic en el botón desplegable (ícono)
                        # Bloque para interactuar con el menú desplegable
                        try:
                            # 1. Clic en el ícono del engranaje
                            icono = celda_accion.find_element(By.XPATH, ".//a[contains(@class, 'dropdown-toggle')]")
                            driver.execute_script("arguments[0].scrollIntoView(true);", icono)
                            ActionChains(driver).move_to_element(icono).click().perform()
                            print("🖱️ Icono desplegable cliceado.")
                            time.sleep(2)  # espera a que aparezca la ventana

                            # 2. Forzar visualización del menú
                            menu = celda_accion.find_element(By.XPATH, ".//ul[contains(@class, 'dropdown-menu')]")
                            driver.execute_script("arguments[0].style.display = 'block';", menu)
                            print("✅ Menú desplegable forzado a visible.")
                            time.sleep(1)

                            #---Haciendo dinamico el encontrar el ID
                            # Encuentra TODOS los enlaces <a> en el menú desplegable
                            links = menu.find_elements(By.TAG_NAME, "a")
                            boton_descarga = None

                            for link in links:
                                texto = link.text.strip()
                                titulo = link.get_attribute("title")
                                id_link = link.get_attribute("id")
                                #print(f"Probando link: texto='{texto}', title='{titulo}', id='{id_link}'")
                                # Busca por texto, título o clase
                                if "Descarga" in texto or "Descarga" in (titulo or ""):
                                    boton_descarga = link
                                    break

                            # Si aún no lo encuentras, podrías tomar el primer <a>
                            if boton_descarga is None and links:
                                print("❌ No se encontró 'Descarga' en texto/título, se usará el primer <a> del menú.")
                                boton_descarga = links[0]

                            # ------------------------------- fin 
                            wait.until(EC.visibility_of(boton_descarga))

                            destino_factura = f"{ruta_carpeta_facturas}/{numero_poliza}_{comprobante_valor}"

                            # Siempre guarda tu ventana principal (X)
                            ventana_cia = driver.current_window_handle

                            # 4. Descargar y continuar el flujo
                            if click_descarga_opcion(driver, destino_factura, boton_descarga,numero_poliza,ruta_carpeta_errores):
                                
                                time.sleep(3)
                                ruta_factura = os.path.join(ruta_carpeta_facturas, f"{numero_poliza}_{comprobante_valor}.pdf")

                                #print(f"📂 Archivo a utilizar:{ruta_factura}")

                                # Lista para guardar las fechas
                                fechas_habiles = []

                                fecha_emision_probar = datetime.strptime(fecha_emision_valor, "%d/%m/%Y")

                                # Agregamos la fecha inicial sí o sí
                                fechas_habiles.append(fecha_emision_probar.strftime("%d/%m/%Y"))

                                # Hasta tener 15 fechas consecutivas (incluye sábados y domingos)
                                while len(fechas_habiles) < 15:
                                    fecha_emision_probar += timedelta(days=1)

                                    # Si la siguiente fecha es mayor que hoy, se detiene
                                    if fecha_emision_probar.date() >= get_fecha_hoy().date():
                                        break

                                    fechas_habiles.append(fecha_emision_probar.strftime("%d/%m/%Y"))

                                # Mostramos las fechas para probar, se considera la fecha de emision de la compañias hasta 7 dia mas, no incluye sabados ni domingos
                                print(f"📅 Fechas de Emisión a probar:{fechas_habiles}")

                                for fecha in fechas_habiles:
                                    print("---------------------------------------")
                                    print(f"⌛ Probando con la Fecha hábil: {fecha}")

                                    #------------INGRESA A SUNAT-------  
                                    nombre_imagen_sunat = f"{numero_proforma}_{numero_poliza}.png"
                                    ruta_imagen_sunat = os.path.join(ruta_carpeta_comprobante, nombre_imagen_sunat)
                                    resultado = consultarValidezSunat(driver,wait,ruc_compania,tipo_doc_birlik,numero_ruc,comprobante_valor,fecha,importe_valor,ruta_imagen_sunat)

                                    driver.switch_to.window(ventana_cia)
                                    print("🔄 Volviendo a la ventana de la CIA")

                                    if resultado is None:
                                        resultado_accion = f'=HYPERLINK("{login_sunat}", "Sunat Bloqueado")'
                                        break
                                    elif resultado:

                                        resultado_sunat = True

                                        if estadoCuota_birlik == "Pendiente-comprobante":
                                            print("📤 Subiendo comprobante a Birlik...")
                                            agregar_comprobante_pago(driver,wait,id_cuota,ruta_factura)
                                            resultado_accion = "Factura Enviada Anteriormente"
                                        else:
                                            # Subir comprobante a Birlik
                                            print("📤 Subiendo todos los documentos a Birlik...")
                                            cancelar_y_agregar_cuota(driver,wait,id_cuota,comprobante_valor,fecha,ruta_factura,ruta_imagen_sunat,resultado_importe)
                                            resultado_accion = f'=HYPERLINK("{url_cuotas_canceladas}{fk_Cliente}", "Enviar Factura")'
            
                                        resultado_birlik = True
                                        break  # Salir del bucle porque ya funcionó con esa fecha

                                    else:
                                        resultado_accion = f'=HYPERLINK("{login_sunat}", "Ver Sunat")'
                                        continue # Si no es True, salta al siguiente intento
                            else:
                                print("❌ No se logró descargar correctamente el archivo PDF.")
                                driver.save_screenshot(f"{ruta_carpeta_errores}/{numero_proforma}_{numero_poliza}.png")

                        except Exception as ex:
                            print(f"Error general al intentar interactuar con el menú desplegable, Detalles {ex}")
    
            if not fila_encontrada_descarga:
                print("❌ No se encontró la fila con el Documento esperado,⌛ Consultando en la otra Compañia...")
                resultado_sunat_cre, resultado_birlik_cre ,resultado_ocr_cre ,resultado_accion_cre = buscaryRegistrarenCrecer(driver,wait,fecha_emision_valor,comprobante_valor,importe_valor,id_cuota,ruc_compania,numero_ruc,numero_proforma,numero_poliza,estadoCuota_birlik,fk_Cliente,tipo_doc_birlik,ruta_carpeta_facturas,ruta_carpeta_comprobante,ruta_carpeta_errores,resultado_importe)
                resultado_sunat = resultado_sunat_cre
                resultado_birlik = resultado_birlik_cre
                resultado_ocr = resultado_ocr_cre
                resultado_accion = resultado_accion_cre

    except Exception as e:
        print(f"⚠️ Detalles del error: {e}")
    finally:
        print("✅ Flujo completado correctamente para la fila.")
        driver.quit()
        return f"Coinciden" if resultado_importe else f"No coinciden" ,"Válido" if resultado_sunat else "No Válido" ,"Cuota Cancelada" if resultado_birlik else "Cuota Pendiente" ,"Si" if resultado_ocr else "No",resultado_estado,resultado_accion

def main():
    
    #------API---------

    todos_los_datos = consultarAPI(url_datos_para_cancelar_cuotas,ids_compania)

    if not todos_los_datos:

        print("❌ No se recibió información de ninguna compañía.")

    else:

        log_path,ruta_salida_API,ruta_salida,ruta_maestro,nombre_hoja, ruta_carpeta_facturas, ruta_carpeta_comprobante, ruta_carpeta_errores,carpeta_compañia,carpeta_principal = crearCarpetas(todos_los_datos,nombre_carpeta_compañia,tipo=2,cia_a_verificar=None)

        try:
            print("\n📁 Iniciando procesamiento para Sanitas Protecta y Crecer...")

            try:
                df = pd.read_excel(ruta_salida_API, engine="openpyxl",dtype={"numeroDocumento": str})
            except Exception as e:
                raise Exception(f" Error al leer el archivo Excel: {e}")
            
            # Nuevas columnas para registrar los resultados
            df["Importe"] = ""
            df["Sunat"] = ""
            df["Birlik"] = ""
            df["OCR"] = ""
            df["Estado"] = ""
            df["Acción"] = ""

            total_filas = len(df)

            for index, row in df.iterrows():
                print(f"\n--- Procesando fila {index + 2} de {total_filas + 1} ---")

                try:

                    importe_estado, sunat_estado, birlik_estado, fecha_detectada,estado_estado, accion_estado = procesar_fila(
                        row,ruta_carpeta_facturas, ruta_carpeta_comprobante, ruta_carpeta_errores,carpeta_compañia)
                    df.at[index, "Importe"] = importe_estado
                    df.at[index, "Sunat"] = sunat_estado
                    df.at[index, "Birlik"] = birlik_estado
                    df.at[index, "OCR"] = fecha_detectada
                    df.at[index, "Estado"] = estado_estado
                    df.at[index, "Acción"] = accion_estado

                    df.to_excel(ruta_salida, index=False)

                    print(f"✅ Fila {index} guardada correctamente.")

                except Exception as e:
                    print(f"❌ Error procesando fila {index}: {e}")
                    df.to_excel(ruta_salida, index=False)
                    continue  
                
                time.sleep(2)
            
            # Guardar archivo original con los resultados
            #df.to_excel(ruta_salida, index=False)
            
            # Guardar el DataFrame en un nuevo archivo de forma individual (ruta_salida),Normalmente es Sheet1 en nuevos Excel
            #guardar_excel_con_formato(ruta_salida,'Sheet1')
                
            print(f"\n✅ Flujo finalizado.")

        finally:

            if os.path.exists(carpeta_principal):
                shutil.rmtree(carpeta_compañia)
                print("🧹 Carpeta eliminada correctamente")
            else:
                print("⚠️ La carpeta no existe")

            time.sleep(2)

if __name__ == "__main__":
    main()
