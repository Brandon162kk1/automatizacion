#--- Imports --
import pandas as pd 
import os
import time
import subprocess
import shutil
#--- Froms ---
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from datetime import datetime
from dateutil.relativedelta import relativedelta
from Sunat.validar_factura import consultarValidezSunat,login_sunat
from Birlik.cancelar_cuotas import agregar_comprobante_pago,cancelar_y_agregar_cuota,url_cuotas_canceladas,url_datos_para_cancelar_cuotas
from Apis.api_birlik import consultarAPI
from GoogleChrome.chromeDriver import abrirDriver, crearCarpetas
from GoogleChrome.fecha_y_hora import get_timestamp,get_fecha_actual
from Correo.correo_it import enviarCaptcha

#------------CRECER----------------
ruc_crecer_vly = '20600098633'
ids_compania = [32]
#-------------CREDENCIALES---------------
login_url_crecer_vida_ley = os.getenv("login_url_crecer_vida_ley")
username_crecer = os.getenv("username_crecer")
password_crecer = os.getenv("password_crecer")
para_venv = os.getenv("para")
para_lista = para_venv.split(",") if para_venv else []
copia_venv = os.getenv("copia")
copias_lista = copia_venv.split(",") if copia_venv else []
#----- Carpeta de la Compañia -------
nombre_carpeta_compañia = f"Crecer_VidaLey_{get_timestamp()}"
# --- Configuración 2Captcha---
API_KEY = "8b61feec172173ef48060a723af1b6c7"

#------- Errores comunes de 2Captcha
# ERROR_WRONG_USER_KEY → API key incorrecta.
# ERROR_KEY_DOES_NOT_EXIST → API key no válida.
# ERROR_ZERO_BALANCE → No tienes saldo.
# ERROR_WRONG_GOOGLEKEY → El sitekey no corresponde o no es válido.
# ERROR_PAGEURL → La URL de la página no es válida.

def bloquear_interaccion():
    subprocess.run(["x11vnc", "-remote", "viewonly"], check=False)
    print("🔒 Interacción humana BLOQUEADA (VNC view-only)")

def desbloquear_interaccion():
    subprocess.run(["x11vnc", "-remote", "noviewonly"], check=False)
    print("✋ Interacción humana HABILITADA")

def click_descarga_opcion(driver,wait,destino_factura, boton_descarga):
    try:

        driver.execute_script("arguments[0].click();", boton_descarga)
        print("✅ Se hizo clic con JS en el botón de descarga.")

        # Esperar a que el loader desaparezca
        wait.until(EC.invisibility_of_element_located((By.XPATH, "//ngx-spinner")))

        time.sleep(3)
        print("⌛ Esperando la ventana descarga de Linux Debian...")
        subprocess.run(["xdotool", "search", "--name", "Save File", "windowactivate", "windowfocus"])
        print("Se hizo FOCO en la nueva ventana de dialogo de Linux Debian")
        time.sleep(5)     
        subprocess.run(["xdotool", "type","--delay", "100", destino_factura])
        print("📄 Se escribió el nombre del archivo")
        time.sleep(3)
        subprocess.run(["xdotool", "key", "Return"])
        print("🖱️ Se presionó Enter para confirmar la descarga.")
        time.sleep(2)
        return True

    except Exception as ex:
        print(f"❌ Error durante el flujo de descarga: {ex}")
        return False

def procesar_fila(driver,wait,row,ruta_carpeta_facturas,ruta_carpeta_comprobante,ruta_carpeta_errores,carpeta_compañia):

    # Extraer valores y quitar espacios en blanco
    id_cuota_birlik = str(row["id_Cuota"]).strip()
    ruc_cliente_birlik = str(row["numeroDocumento"]).strip()
    tipo_doc_birlik = str(row["tipoDocumento"]).strip()
    numero_proforma_birlik = str(row["codigoCuota"]).strip()
    numero_poliza_birlik = str(row["numeroPoliza"]).strip()
    importe_total_birlik = str(row["importe"]).strip()
    estado_Cuota_birlik = str(row["estadoCuota"]).strip()
    fk_Cliente_birlik = str(row["fk_Cliente"]).strip() 
    vigencia_inicio_birlik = str(row["vigenciaInicio"]).strip()

    # Convertir a objeto datetime
    fecha_inicio_dt = datetime.strptime(vigencia_inicio_birlik, "%d/%m/%Y")
    fecha_fin_dt = datetime.strptime(vigencia_inicio_birlik,"%d/%m/%Y")
    # Restar 1 mes
    fecha_menos_un_mes = fecha_inicio_dt - relativedelta(months=1)
    fecha_mas_un_mes = fecha_fin_dt + relativedelta(months=1)
    # Convertir de nuevo a string si lo necesitas en ese formato
    fecha_resultado_menos = fecha_menos_un_mes.strftime("%d/%m/%Y")
    fecha_resultado_mas = fecha_mas_un_mes.strftime("%d/%m/%Y")
   
    resultado_importe = False
    resultado_sunat = False
    resultado_birlik = False
    resultado_estado = None
    resultado_accion = ""

    try:
        
        input_fullname = wait.until(EC.visibility_of_element_located((By.ID, "sfullname")))
        input_fullname.clear()
        input_fullname.send_keys(ruc_cliente_birlik)
        print(f"✅ Se ingresó el RUC '{ruc_cliente_birlik}'.")

        # Quitar el atributo readonly con Javascript
        input_fecha = driver.find_element(By.ID, "solicitud_fecha_inicio")
        driver.execute_script("arguments[0].removeAttribute('readonly')", input_fecha)
        input_fecha.clear()
        input_fecha.send_keys(fecha_resultado_menos)
        print("✅ Se ingresó un mes antes de la fecha de inicio de la vigencia")

        time.sleep(3)

        # Quitar el atributo readonly con Javascript
        input_fecha_fin = driver.find_element(By.ID, "solicitud_fecha_fin")
        driver.execute_script("arguments[0].removeAttribute('readonly')", input_fecha_fin)
        input_fecha_fin.clear()
        input_fecha_fin.send_keys(fecha_resultado_mas)
        print("✅ Se ingresó un mes despues de la fecha fin de la vigencia")

        input_poliza = wait.until(EC.visibility_of_element_located((By.XPATH,"//label[normalize-space()='Póliza']/following-sibling::input")))
        input_poliza.clear()
        input_poliza.send_keys(numero_poliza_birlik)
        print(f"✅ Se ingresó la Póliza '{numero_poliza_birlik}'.")

        # Esperar que el select esté visible
        select_element = wait.until(EC.visibility_of_element_located((By.ID, "Estados")))
        select = Select(select_element)
        select.select_by_value("0")
        print("🖱️ Se selecciono la opcion 'Todos' ")

        # Esperar que el botón esté visible y clickable
        boton_consultar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Consultar')]")))
        # Click en el botón usando JavaScript (más efectivo en Angular/React)
        driver.execute_script("arguments[0].click();", boton_consultar)
        print("🖱️ Clic en Consultar")
        time.sleep(5)

        # Esperar a que el loader desaparezca
        wait.until(EC.invisibility_of_element_located((By.XPATH, "//ngx-spinner")))

        # Hacer click en un lugar vacío de la página (por ejemplo el body)
        body = driver.find_element(By.TAG_NAME, "body")
        ActionChains(driver).move_to_element(body).click().perform()
        print("🖱️ Clic en un lugar vacío de la página para asegurar la interacción.")

        # Seleccionar la opción "20" en el dropdown de cantidad de registros
        select_elem = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "select.select2.form-control.input-sm")))
        Select(select_elem).select_by_value("20")
        print("✅ Se seleccionó '20' registros para mostrar más filas.")
        time.sleep(5)

        # Esperar a que el loader desaparezca
        wait.until(EC.invisibility_of_element_located((By.XPATH, "//ngx-spinner")))

        try:
            table = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.table.table-striped.table-bordered")))
            rows = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "table.table.table-striped.table-bordered tbody tr")))
            print(f"✅ ¡Tabla cargada con {len(rows)} filas!")

        except TimeoutException:
            driver.save_screenshot(os.path.join(ruta_carpeta_errores,f"{numero_poliza_birlik}_{get_timestamp()}.png"))
            raise Exception("❌ La tabla no tiene filas")

        fila_encontrada = False

        print("-------------------------------------------------")
        num_poliza_mas_reno = f"{numero_poliza_birlik}-R{numero_proforma_birlik}"
        print(f"🔍 Buscando el código de cuota '{numero_proforma_birlik}' o el numero de poliza '{num_poliza_mas_reno}'.")

        for row in rows:

            cells = row.find_elements(By.TAG_NAME, "td")

            if len(cells) < 15:
                continue

            cod_cotizacion = cells[1].text.strip()                                  # Código Cotización
            prima_total = cells[4].text.replace("S/.", "").replace(",", "").strip() # Monto Total
            estado_compania = cells[5].text.strip()                                 # "Pendiente" "Aprobado" "Rechazado"
            estado_emision = cells[11].text.strip()                                 # "Pendiente" "En Proceso" "Realizado" "Anulado"
            numero_poliza_mas_reno = cells[9].text.strip()                          # Numero Poliza 
            fecha_emision = cells[12].text.strip()                                  # Fecha emision PDF
            estado_pago = cells [16].text.strip()                                   # Estado de pago: "Pagado" 
            comprobante = cells[17].text.strip()                                    # Numero de Comprobante
            fecha_emision_comprobante =cells[18].text.strip()                       # Fecha de Emision Comprobante

            if cod_cotizacion == numero_proforma_birlik or num_poliza_mas_reno == numero_poliza_mas_reno :

                fila_encontrada = True
                print("✅ Fila encontrada.")
                resultado_estado = estado_pago

                diferencia = abs(float(prima_total) - float(importe_total_birlik))
                print(f"Importe de Birlik: {float(importe_total_birlik)} -- Importe de la Compañía : {float(prima_total)}")
                if diferencia > 0.05:
                    print("❌ Los importes No coinciden")
                else:
                    print("✅ Los importes Coinciden")
                    resultado_importe = True

                if estado_pago.lower() == "pagado":

                    if len(cells) > 18:

                        comprobante_pdf = cells[17]

                        try:
                            # Hacer clic en el link de esa celda
                            link = comprobante_pdf.find_element(By.TAG_NAME, "a")

                            ventana_principal_crecer = driver.current_window_handle

                            destino_factura = f"{ruta_carpeta_facturas}/{numero_poliza_birlik}_{comprobante}"

                            for i in range(2):  # aparecen 2 archivos
    
                                if i == 0:
                                    print("📥 Descargando Factura (1/2)...")

                                    if click_descarga_opcion(driver,wait, destino_factura, link):
                                        print("✅ Factura descargada")

                                    time.sleep(2)

                                else:
                                    print(f"⌛ Cancelando descarga automática {i+1}/2...")
                                    subprocess.run(["xdotool", "key", "Escape"])
                                    print("⌨️ Se presionó ESC para cancelar la descarga.")
                                    time.sleep(3)

                            # Esperar a que el loader desaparezca
                            wait.until(EC.invisibility_of_element_located((By.XPATH, "//ngx-spinner")))

                            ruta_factura = os.path.join(ruta_carpeta_facturas, f"{numero_poliza_birlik}_{comprobante}.pdf")

                            if os.path.exists(ruta_factura):
                                #------------INGRESA A SUNAT-------  
                                nombre_imagen_sunat = f"{numero_proforma_birlik}_{numero_poliza_birlik}.png"
                                ruta_imagen_sunat = os.path.join(ruta_carpeta_comprobante, nombre_imagen_sunat)
                                resultado = consultarValidezSunat(driver,wait,ruc_crecer_vly,tipo_doc_birlik,ruc_cliente_birlik,comprobante,fecha_emision_comprobante,prima_total,ruta_imagen_sunat)

                                driver.switch_to.window(ventana_principal_crecer)
                                print("🔄 Volviendo a la ventana de la CIA")

                                if resultado is None:
                                        resultado_accion = f'=HYPERLINK("{login_sunat}", "Sunat Bloqueado")'
                                        break
                                elif resultado:

                                    resultado_sunat = True

                                    if estado_Cuota_birlik == "Pendiente-comprobante":

                                        print("📤 Subiendo comprobante a Birlik...")
                                        agregar_comprobante_pago(driver,wait,id_cuota_birlik,ruta_factura)
                                        resultado_accion = "Factura Enviada Anteriormente"

                                    else:

                                        print("📤 Subiendo todos los documentos a Birlik...")
                                        cancelar_y_agregar_cuota(driver,wait,id_cuota_birlik,comprobante,fecha_emision,ruta_factura,ruta_imagen_sunat,resultado_importe)
                                        resultado_accion = f'=HYPERLINK("{url_cuotas_canceladas}{fk_Cliente_birlik}", "Enviar Factura")'
            
                                    resultado_birlik = True
                                    break
                                else:
                                    resultado_accion = f'=HYPERLINK("{login_sunat}", "Ver Sunat")'
                                                                          
                            else:
                                raise Exception(f"No se descargo la Factuta '{comprobante}'.")

                        except Exception as e:
                            print(f"No se pudo descargar el PDF,Flujo Terminado con estado {estado_pago}, error: {e}")

                else:
                    print(f"La Poliza {numero_poliza_birlik} tiene estado de Pago '{estado_pago}',estado de emision '{estado_emision}', y estado de compañia '{estado_compania}'.")
                    resultado_accion  = "Esperar a que pague"

        if not fila_encontrada:
            resultado_accion = "°PROFORMA INCORRECTO"
            resultado_estado  = "No se sabe"
            print(f"❌ No se encontró ninguna fila que coincida con el código de cuota {numero_proforma_birlik} o {numero_poliza_birlik}-R{numero_proforma_birlik}.")
        
    except Exception as e:
        print(f"Detalles del error: {e}")
    finally:
        driver.refresh()
        return f"Coinciden" if resultado_importe else f"No coinciden" ,"Válido" if resultado_sunat else "No Válido" ,"Registrado" if resultado_birlik else "No registrado" , resultado_estado, resultado_accion
    
def main():
    
    puerto = os.getenv("NOVNC_PORT")
    display_num = os.getenv("DISPLAY_NUM", "0")
    os.environ["DISPLAY"] = f":{display_num}"

    driver, wait = abrirDriver()

    driver.get(login_url_crecer_vida_ley)
    print("⌛ Ingresando a la URL")

    user_input = wait.until(EC.presence_of_element_located((By.ID, "suser")))
    user_input.clear()
    user_input.send_keys(username_crecer)
    print("⌨️ Digitando el Username")
        
    pass_input = wait.until(EC.presence_of_element_located((By.ID, "spassword")))
    pass_input.clear()
    pass_input.send_keys(password_crecer)
    print("⌨️ Digitando el Password")
   
    print("🧩 Resuelve el CAPTCHA manualmente y clic en 'Ingresar'.")

    #------Carpetas de Descargas y Volumen del Docker----------
    carpeta_descargas = "Downloads"
    # --- Construir ruta de Downloads por defecto ---
    base_dir = os.path.dirname(os.path.abspath(__file__))
    ruta_carpeta_downloads = os.path.join(base_dir,carpeta_descargas)
    #------ Carpeta Principal -------
    nom_carp_principal= f"Reporte_Cuotas_Diarias_{get_fecha_actual()}"
    carpeta_principal = os.path.join(ruta_carpeta_downloads, nom_carp_principal)
    # Esto crea la carpeta si no existe
    os.makedirs(carpeta_principal, exist_ok=True) 
    ruta_imagen = os.path.join(carpeta_principal,f"{get_timestamp()}.png")
    driver.save_screenshot(ruta_imagen)
    enviarCaptcha(para_lista,copias_lista,puerto,"Crecer Vida Ley",ruta_imagen)
    # Espera humana (hasta 5 minutos)
    wait_humano = WebDriverWait(driver,300)
    wait_humano.until(EC.presence_of_element_located((By.XPATH, "//a[contains(normalize-space(),'Cerrar sesión')]")))

    bloquear_interaccion()

    print("✅ Login exitoso detectado (Cerrar sesión visible)")
    print("🚀 Continuando flujo automáticamente")

    # iframe_tag = driver.find_element(By.XPATH, "//iframe[contains(@src, 'recaptcha')]")
    # sitekey_url = iframe_tag.get_attribute("src")

    # parsed = urlparse.urlparse(sitekey_url)
    # sitekey = urlparse.parse_qs(parsed.query)["k"][0]
    # print(f"✅ Sitekey encontrada: {sitekey}")

    # # Enviar captcha a 2Captcha
    # resp = requests.post("http://2captcha.com/in.php", data={
    #     "key": API_KEY,
    #     "method": "userrecaptcha",
    #     "googlekey": sitekey,
    #     "pageurl": driver.current_url
    # })
    # if not resp.text.startswith("OK|"):
    #     raise Exception(f"❌ Error al enviar captcha a 2Captcha: {resp.text}")

    # captcha_id = resp.text.split('|')[1]
    # print(f"⌛ Captcha enviado a 2Captcha. ID: {captcha_id}")

    # # Esperar la respuesta de 2Captcha
    # result_url = f"http://2captcha.com/res.php?key={API_KEY}&action=get&id={captcha_id}"
    # token = None
    # for i in range(30):  # 30 intentos x 5 seg = máx 150 seg
    #     time.sleep(5)
    #     res = requests.get(result_url)
    #     if "OK" in res.text:
    #         token = res.text.split('|')[1]
    #         print("✅ Captcha resuelto por 2Captcha")
    #         break

    # if not token:
    #     raise Exception("❌ No se pudo resolver el captcha")

    # # Inyectar el token en el campo oculto del captcha
    # driver.execute_script('document.getElementById("g-recaptcha-response").value = arguments[0];', token)

    # ingresar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Ingresar')]")))
    # ingresar_btn.click()
    # print("🖱️ Clic en 'Ingresar' para Crecer Seguros.")
    # time.sleep(5)
           
    # Esperamos que el SVG esté cargado (opcional pero recomendado)
    span_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Gestión de Cotización']")))
    actions = ActionChains(driver)
    actions.move_to_element(span_element).click().perform()
    print("🖱️ Clic en 'Gestion de Cotizacion'")
    time.sleep(3)

    # Esperar que el enlace esté disponible
    link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Bandeja de Evaluación")))
    link.click()
    print("🖱️ Clic en 'Bandeja de Evaluación'")

    time.sleep(3)

    while True:

        try:
            todos_los_datos = consultarAPI(url_datos_para_cancelar_cuotas,ids_compania)

            if not todos_los_datos:
                raise Exception("No se recibió información de ninguna compañía.")

            log_path,ruta_salida_API,ruta_salida,ruta_maestro,nombre_hoja, ruta_carpeta_facturas, ruta_carpeta_comprobante, ruta_carpeta_errores,carpeta_compañia,carpeta_principal = crearCarpetas(todos_los_datos,nombre_carpeta_compañia,tipo=2,cia_a_verificar=None)

            print("\n📁 Iniciando procesamiento para Crecer Vida Ley...")

            try:
                df = pd.read_excel(ruta_salida_API, engine="openpyxl",dtype={"numeroDocumento": str})
            except Exception as e:
                raise Exception(f" Error al leer el archivo Excel: {e}")

            df["Importe"] = ""
            df["Sunat"] = ""
            df["Birlik"] = ""
            df["Estado"] = ""
            df["Acción"] = ""

            total_filas = len(df)

            for index, row in df.iterrows():
                print(f"\n--- Procesando fila {index + 2} de {total_filas + 1} ---")

                try:

                    importe_estado,sunat_estado,birlik_estado,estado_estado,accion_estado = procesar_fila(
                        driver,wait,row,ruta_carpeta_facturas,ruta_carpeta_comprobante,ruta_carpeta_errores,carpeta_compañia)

                    df.at[index, "Importe"] = importe_estado
                    df.at[index, "Sunat"] = sunat_estado
                    df.at[index, "Birlik"] = birlik_estado
                    df.at[index, "Estado"] = estado_estado
                    df.at[index, "Acción"] = accion_estado

                    df.to_excel(ruta_salida, index=False)

                    print(f"✔ Fila {index} guardada correctamente.")

                except Exception as e:
                    print(f"❌ Error procesando fila {index}: {e}")
                    df.to_excel(ruta_salida, index=False)
                    continue  
                
                time.sleep(3)
            
            df.to_excel(ruta_salida, index=False)

            #guardar_excel_con_formato(ruta_salida,'Sheet1')

        except Exception as e:
            print(f"\n🟡 Proceso Detenido, por: {e}")
        finally:
            if os.path.exists(carpeta_principal):
                shutil.rmtree(carpeta_compañia)
                print("🧹 Carpeta eliminada correctamente")
            else:
                print("⚠️ La carpeta no existe")
            print(f"\n✅ Flujo finalizado, Intentando de nuevo en 10 segundos.")
            time.sleep(10)
   
if __name__ == "__main__":
    main()