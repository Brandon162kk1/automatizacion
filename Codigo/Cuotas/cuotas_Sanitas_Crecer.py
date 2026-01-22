#-- Imports ---
import subprocess
import time
import os
#-- Froms ----
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import Select
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from selenium.webdriver import ActionChains
from Sunat.validar_factura import consultarValidezSunat,login_sunat
from Birlik.cancelar_cuotas import agregar_comprobante_pago, cancelar_y_agregar_cuota,url_cuotas_canceladas
from GoogleChrome.fecha_y_hora import get_fecha_hoy

#---------URL DE SANITAS CRECER -------
login_url_sanitas_crecer = os.getenv("login_url_sanitas_crecer")
#-------CREDENCIALES SANITAS----------
username_sanitas = os.getenv("usernameSanitas")
password_sanitas = os.getenv("passwordSanitas")

def click_descarga_opcion(driver, destino_factura, boton_descarga,numero_poliza,ruta_carpeta_errores):
    try:

        driver.execute_script("arguments[0].click();", boton_descarga)
        print("🖱️ Se hizo clic con JS en el botón de descarga.")
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
        return True

    except Exception as ex:
        print("❌ Error durante el flujo de descarga:", ex)
        driver.save_screenshot(f"{ruta_carpeta_errores}/{numero_poliza}_ventanalinux.png")
        return False

def buscaryRegistrarenCrecer(
driver,
wait,
fecha_emision_valor,
comprobante_valor,
importe_valor,
id_cuota,
ruc_compania,
numero_ruc,
numero_proforma,
numero_poliza,
estadoCuota_birlik,
fk_Cliente,
tipo_doc_birlik,ruta_carpeta_facturas,ruta_carpeta_comprobante,ruta_carpeta_errores,resultado_importe):
    
    resultado_ocr = False
    resultado_sunat = False
    resultado_birlik = False
    resultado_accion = ""

    try:
            driver.get(login_url_sanitas_crecer)
            print("✅ Ingresando a la URL")
        
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
        
            wait.until(EC.url_contains("Quotation/Index"))
        
            autogestion_link = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//a[contains(@class,'dropdown-toggle') and contains(text(),'Autogestión')]")
            ))
            autogestion_link.click()
            print("✅ Se hizo clic en 'Autogestión'.")
       
            consulta_comprobantes_link = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//a[@href='/PaymentVouchers/LegalDocumentsSelfManagementIndex' and contains(text(),'Comprobantes')]")
            ))
            consulta_comprobantes_link.click()
            print("🌐 Navegando a 'Consulta de Comprobantes de pago'.")
            time.sleep(3)

            contract_input = wait.until(EC.presence_of_element_located((By.ID, "ContractNumber")))
            contract_input.clear()
            contract_input.send_keys(numero_poliza)
            print(f"✅ Se ingresó el número de póliza: {numero_poliza}")

            fecha = datetime.strptime(fecha_emision_valor, "%d/%m/%Y")
            fecha_inicio = fecha - relativedelta(months=1)
            fecha_fin = fecha + relativedelta(months=1)

            rango_fechas = f"{fecha_inicio.strftime('%d/%m/%Y')} - {fecha_fin.strftime('%d/%m/%Y')}"
            print("📅 Rango de fechas:", rango_fechas)

            input_rango = wait.until(EC.presence_of_element_located((By.ID, "DatesRangeCreation")))
            input_rango.clear()
            input_rango.send_keys(rango_fechas)

            filter_btn = wait.until(EC.element_to_be_clickable((By.ID, "btnSearch")))
            filter_btn.click()

            select_elem = wait.until(EC.presence_of_element_located((By.NAME, "LegalDocumentSelfManagementListTable_length")))
            Select(select_elem).select_by_value("1000")
            print("🖱️ Seleccionado '1000' registros para ver más filas en la tabla de comprobantes.")
            time.sleep(2)
    
            tabla = wait.until(EC.presence_of_element_located((By.ID, "LegalDocumentSelfManagementListTable")))
            filas_tabla = tabla.find_elements(By.TAG_NAME, "tr")
    
            fila_encontrada_descarga = False
            for fila in filas_tabla:
                celdas = fila.find_elements(By.TAG_NAME, "td")
                if len(celdas) >= 4:
                    doc_valor = celdas[2].text.strip()
                    if comprobante_valor in doc_valor:
                        fila_encontrada_descarga = True
                        print(f"✅ Fila encontrada en 'Consulta de Comprobantes': Documento = {doc_valor}")
                
                        celda_accion = celdas[-1]
                        
                        try:
                            icono = celda_accion.find_element(By.XPATH, ".//a[contains(@class, 'dropdown-toggle')]")
                            driver.execute_script("arguments[0].scrollIntoView(true);", icono)
                            ActionChains(driver).move_to_element(icono).click().perform()
                            print("🖱️ Icono desplegable clickeado.")
                            time.sleep(2)

                            menu = celda_accion.find_element(By.XPATH, ".//ul[contains(@class, 'dropdown-menu')]")
                            driver.execute_script("arguments[0].style.display = 'block';", menu)
                            print("✅ Menú desplegable forzado a visible.")
                            time.sleep(1)

                            links = menu.find_elements(By.TAG_NAME, "a")
                            boton_descarga = None

                            for link in links:
                                texto = link.text.strip()
                                titulo = link.get_attribute("title")
                                id_link = link.get_attribute("id")

                                if "Descarga" in texto or "Descarga" in (titulo or ""):
                                    boton_descarga = link
                                    break

                            if boton_descarga is None and links:
                                print("❌ No se encontró 'Descarga' en texto/título, se usará el primer <a> del menú.")
                                boton_descarga = links[0]

                            wait.until(EC.visibility_of(boton_descarga))

                            destino_factura = f"{ruta_carpeta_facturas}/{numero_poliza}_{comprobante_valor}"

                            ventana_cia = driver.current_window_handle

                            if click_descarga_opcion(driver, destino_factura, boton_descarga,numero_poliza,ruta_carpeta_errores):
                                
                                time.sleep(3)
                                ruta_factura = os.path.join(ruta_carpeta_facturas, f"{numero_poliza}_{comprobante_valor}.pdf")

                                print(f"📂 Archivo a utilizar:{ruta_factura}")

                                fechas_habiles = []

                                fecha_emision_probar = datetime.strptime(fecha_emision_valor, "%d/%m/%Y")

                                fechas_habiles.append(fecha_emision_probar.strftime("%d/%m/%Y"))

                                while len(fechas_habiles) < 15:
                                    fecha_emision_probar += timedelta(days=1)

                                    if fecha_emision_probar.date() >= get_fecha_hoy().date():
                                        break

                                    fechas_habiles.append(fecha_emision_probar.strftime("%d/%m/%Y"))

                                print(f"📅 Fechas de Emisión a probar:{fechas_habiles}")

                                for fecha in fechas_habiles:
                                    print("---------------------------------------")
                                    print(f"⌛ Probando con la Fecha hábil: {fecha}")

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
                                            print("📤 Subiendo todos los documentos a Birlik...")
                                            cancelar_y_agregar_cuota(driver,wait,id_cuota,comprobante_valor,fecha,ruta_factura,ruta_imagen_sunat,resultado_importe)
                                            resultado_accion = f'=HYPERLINK("{url_cuotas_canceladas}{fk_Cliente}", "Enviar Factura")'
            
                                        resultado_birlik = True
                                        break

                                    else:
                                        resultado_accion = f'=HYPERLINK("{login_sunat}", "Ver Sunat")'
                                        continue

                            else:
                                driver.save_screenshot(f"{ruta_carpeta_errores}/{id_cuota}_{numero_poliza}_DescargarPDF.png")
                                print("❌ No se logró descargar correctamente el archivo PDF.")

                        except Exception as ex:
                            print(f"Error general al intentar interactuar con el menú desplegable, Detalles: {ex}")
                            return resultado_sunat , resultado_birlik , resultado_ocr,resultado_accion  
    
            if not fila_encontrada_descarga:
                print("No se encontró la fila con el Documento esperado en la tabla de comprobantes.")
                driver.save_screenshot(f"{ruta_carpeta_errores}/{id_cuota}_{numero_poliza}_NohayComproFactenlaCompania.png")

            print("Flujo completado correctamente para la fila.")

            driver.quit()
            return resultado_sunat , resultado_birlik , resultado_ocr,resultado_accion 
    except Exception as e:
        print(f"Detalles del error: {e}")
        return resultado_sunat,resultado_birlik,resultado_ocr,resultado_accion
    finally:
        driver.quit()
