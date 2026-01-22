#-- Imports ---
import subprocess
import time
import os
import pandas as pd
import socket
import shutil
import zipfile
import re
import pdfplumber
#-- Froms ----
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from Sunat.validar_factura import consultarValidezSunat,login_sunat
from Birlik.cancelar_cuotas import agregar_comprobante_pago,cancelar_y_agregar_cuota,url_cuotas_canceladas,url_datos_para_cancelar_cuotas
from Apis.api_birlik import consultarAPI
from GoogleChrome.chromeDriver import abrirDriver, crearCarpetas
from GoogleChrome.fecha_y_hora import get_timestamp
from datetime import datetime

#--------- COMPAÑÍA SANITAS QUALITAS------
ruc_qualitas = '20553157014'
# Lista de IDs de compañía
ids_compania = [26]
#-------CREDENCIALES QUALITAS----------
login_url_qualitas = os.getenv("login_url_qualitas")
claveCorredor = os.getenv("claveCorredor")
username = os.getenv("usernameQualitas")
password = os.getenv("passwordQualitas")
nombre = os.getenv("CONT_NAME", socket.gethostname())
#--- NOMBRE SERVICE DNS DOCKER PARA UTILIZAR EN LA API -----
nom_serv = os.getenv("nom_serv")
#----- Carpeta de la Compañia -------
nombre_carpeta_compañia = f"Qualitas_{get_timestamp()}"

def click_descarga_factura(driver, destino_factura, boton_descarga,numero_poliza,ruta_carpeta_errores):
    
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
        print(f"❌ Error durante el flujo de descarga: {ex}")
        driver.save_screenshot(f"{ruta_carpeta_errores}/{numero_poliza}_ventanalinux.png")
        return False

def extraer_datos_pdf(ruta_pdf):
    texto = ""
    with pdfplumber.open(ruta_pdf) as pdf:
        for page in pdf.pages:
            texto += page.extract_text() + "\n"

    # Regex exactos según tu PDF
    regex_fecha = r"Fecha\s+de\s+Emisión\s*:\s*([0-9]{2}-[0-9]{2}-[0-9]{4})"
    regex_comprobante = r"\b([A-Z]{2}[0-9]{2}-[0-9]{6})\b"

    fecha_match = re.search(regex_fecha, texto)
    comprobante_match = re.search(regex_comprobante, texto)

    fecha_emision = fecha_match.group(1) if fecha_match else None
    numero_comprobante = comprobante_match.group(1) if comprobante_match else None

    return fecha_emision, numero_comprobante

def procesar_fila(driver,wait,row,ruta_carpeta_facturas, ruta_carpeta_comprobante, ruta_carpeta_errores,carpeta_compañia):

    #--Extraer valores y quitar espacios en blanco
    numero_poliza_birlik = str(row["numeroPoliza"]).strip()
    tipo_doc_birlik = str(row["tipoDocumento"]).strip()
    ruc_cliente_birlik = str(row["numeroDocumento"]).strip()
    id_cuota_birlik = str(row["id_Cuota"]).strip()
    fk_Cliente_birlik = str(row["fk_Cliente"]).strip() 
    fk_compania_birlik = str(row["fK_Compania"]).strip()
    fk_Ramo_birlik = int(str(row["fk_Ramo"]).strip())
    numero_proforma_birlik = str(row["codigoCuota"]).strip()
    importe_total_birlik = str(row["importe"]).strip()
    estadoCuota_birlik = str(row["estadoCuota"]).strip()
    primaneta_birlik = str(row["primaNeta"]).strip()
    id_Poliza_birlik = str(row["id_Poliza"]).strip()
    fecha_inicioVig_Birlik = str(row["vigenciaInicio"]).strip()
    fecha_finVig_Birlik = str(row["vigenciaFin"]).strip()
    #--------------

    resultado_importe = False
    resultado_sunat = False
    resultado_birlik = False
    resultado_estado = None
    resultado_accion = ""

    # if numero_proforma_birlik != "3061537":
    #     return f"Pagina Web en Mantenimiento" if resultado_importe else "Pagina Web en Mantenimiento" ,"Pagina Web en Mantenimiento" if resultado_sunat else "Pagina Web en Mantenimiento" ,"Pagina Web en Mantenimiento" if resultado_birlik else "Pagina Web en Mantenimiento" , "" if resultado_estado else "Pagina Web en Mantenimiento", "" if resultado_accion else "Pagina Web en Mantenimiento"

    try:

        time.sleep(2)
        # ------------------ Inicio del Flujo de Automatización ------------------
        
        poliza_input = wait.until(EC.visibility_of_element_located((By.ID, "numberPolicy")))
        poliza_input.clear()
        poliza_input.send_keys(numero_poliza_birlik)
        print(f"✅ Póliza ingresada: {numero_poliza_birlik}")

        # Esperar que desaparezca el loader
        wait.until(EC.invisibility_of_element_located((By.ID, "loader")))

        # Esperar botón clickeable
        boton = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Consulta poliza']")))
        boton.click()
        print("🖱️ Clic en 'Buscar'.")

        # Esperar que desaparezca el loader
        wait.until(EC.invisibility_of_element_located((By.ID, "loader")))

        header = wait.until(EC.element_to_be_clickable((By.XPATH, "//p[contains(text(),'Recibos de póliza')]")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", header)
        header.click()
        print("🖱️ Clic en el div de 'Recibos de Poliza'.")
   
        contenedor = wait.until(EC.presence_of_element_located((By.ID, "data-receipts")))

        tabla = contenedor.find_element(By.TAG_NAME, "table")

        filas = tabla.find_elements(By.CSS_SELECTOR, "tbody tr")
        print("Total de filas encontradas:", len(filas))

        encontrado = False

        for i, fila in enumerate(filas, start=1):

            celdas = fila.find_elements(By.CSS_SELECTOR, "th, td")

            if len(celdas) < 3:
                continue

            cols = [c.text.strip().replace("\n", " ") for c in celdas]

            # Variables individuales (no listas)
            codigoCuota = cols[0]                        # Columna 5
            importe = cols[5].replace("$", "").strip()   # Columna 6
            estado = cols[6]                             # Columna 8

            if numero_proforma_birlik in codigoCuota:
                print(f"Fila {i} --> CodigoCuota: {codigoCuota} | Importe: {importe} | Estado: {estado}")
                encontrado = True

                diferencia = abs(float(importe) - float(importe_total_birlik))
                print(f"Importe de Birlik: {float(importe_total_birlik)} -- Importe de la Compañía : {float(importe)}")
                if diferencia > 0.05:
                    print("❌ Los importes No coinciden")
                else:
                    print("✅ Los importes Coinciden")
                    resultado_importe = True
                
                if estado.lower() == 'pagado':

                    destino_carpeta_zip = f"{ruta_carpeta_facturas}/{numero_poliza_birlik}_{numero_proforma_birlik}"
                    ruta_factura = os.path.join(ruta_carpeta_facturas, f"{numero_poliza_birlik}_{numero_proforma_birlik}.pdf")
                    ventana_principal_qualitas = driver.current_window_handle
                    resultado_estado = estado

                    try:

                        boton_kebab = fila.find_element(By.CSS_SELECTOR, ".dropdown-toggle .icon")
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", boton_kebab)
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", boton_kebab)
                        print(f"✅ Kebab menu clickeado en la fila {i}")

                        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".dropdown-menu.show")))

                        dropdown = fila.find_element(By.CSS_SELECTOR, ".dropdown-menu.show")
                        btn_descargar = dropdown.find_element(By.XPATH, ".//button[contains(text(), 'Descargar recibo')]")

                        if click_descarga_factura(driver, destino_carpeta_zip,btn_descargar,numero_poliza_birlik,ruta_carpeta_errores):
                            print("✅ Archivo '.zip' descargado.")
                        else:
                            raise Exception("No se pudo descargar.")

                        time.sleep(2)  

                        nombre_zip = f"{numero_poliza_birlik}_{numero_proforma_birlik}.zip"
                        ruta_zip = os.path.join(ruta_carpeta_facturas, nombre_zip)

                        try:
                            if not os.path.exists(ruta_zip):
                                raise Exception(f"No existe el ZIP descargado: {ruta_zip}")

                            with zipfile.ZipFile(ruta_zip, 'r') as z:
                                pdfs = [f for f in z.namelist() if f.endswith(".pdf")]
                                if not pdfs:
                                    raise Exception("El ZIP no contiene PDF")

                                pdf_interno = pdfs[0]

                                nuevo_nombre_pdf = f"{numero_poliza_birlik}_{numero_proforma_birlik}.pdf"

                                ruta_pdf_salida = os.path.join(ruta_carpeta_facturas, nuevo_nombre_pdf)

                                with open(ruta_pdf_salida, "wb") as f:
                                    f.write(z.read(pdf_interno))

                            print(f"✅ PDF extraído en: {ruta_pdf_salida}")

                        finally:
                            if os.path.exists(ruta_zip):
                                os.remove(ruta_zip)
                                print("🧹 Archivo ZIP eliminado correctamente")
                            else:
                                print("⚠️ El archivo ZIP no existe")

                    except Exception as e:
                        raise Exception(f"Error descargando ZIP, Motivo -> {e}")
                       
                    fecha_emision_pdf,numero_comprobante = extraer_datos_pdf(ruta_pdf_salida)
                    fecha_convertida = fecha_emision_pdf.replace("-", "/")

                    fecha_habiles_factura = []

                    fecha_emision_probar = datetime.strptime(fecha_convertida, "%d/%m/%Y")

                    fecha_habiles_factura.append(fecha_emision_probar.strftime("%d/%m/%Y"))

                    print(f"📅 Fechas de Emisión a probar:{fecha_habiles_factura}")

                    for fecha in fecha_habiles_factura:
                        print("---------------------------------------")
                        print(f"⌛ Probando con la Fecha hábil: {fecha}")

                        nombre_imagen_sunat = f"{numero_proforma_birlik}_{numero_poliza_birlik}.png"
                        ruta_imagen_sunat = os.path.join(ruta_carpeta_comprobante, nombre_imagen_sunat)
                        resultado = consultarValidezSunat(driver,wait,ruc_qualitas,tipo_doc_birlik,ruc_cliente_birlik,numero_comprobante,fecha,importe,ruta_imagen_sunat)

                        driver.switch_to.window(ventana_principal_qualitas)
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
                                cancelar_y_agregar_cuota(driver,wait,id_cuota_birlik,numero_comprobante,fecha,ruta_factura,ruta_imagen_sunat,resultado_importe)
                                resultado_accion = f'=HYPERLINK("{url_cuotas_canceladas}{fk_Cliente_birlik}", "Enviar Factura")'
            
                            resultado_birlik = True
                            break

                        else:
                            resultado_accion = f'=HYPERLINK("{login_sunat}", "Ver Sunat")'
                            continue

                else:
                    resultado_estado = "Por Pagar"
                    resultado_accion = 'Sin Observacion'

                break

        if not encontrado:
            resultado_estado = "No se encuentro cuota"
            resultado_accion = f'=HYPERLINK("{url_cuotas_canceladas}", "Revisar Cuota")'
            raise Exception(f"No se encontró ninguna fila con el código {numero_proforma_birlik}")

    except Exception as e:
        print(f"❌ Error Procesando toda la fila, Motivo: {e}")
    finally:
        return f"Coinciden" if resultado_importe else f"No coinciden" ,"Válido" if resultado_sunat else "No Válido" ,"Cuota Cancelada" if resultado_birlik else "Cuota Pendiente" , "No indica" if resultado_estado is None else resultado_estado ,resultado_accion

def main():
    
    display_num = os.getenv("DISPLAY_NUM", "0")
    os.environ["DISPLAY"] = f":{display_num}"

    driver, wait = abrirDriver()#carpeta_compañia

    # ------------------ Inicio del Flujo de Automatización ------------------
    driver.get(login_url_qualitas) 
    print("⌛ Ingresando a la URL")

    clave_input = wait.until(EC.presence_of_element_located((By.ID, "_com_liferay_login_web_portlet_LoginPortlet_login")))
    clave_input.clear()
    clave_input.send_keys(claveCorredor)
    print("⌨️ Digitando la Clave de Corredor")

    user_input = wait.until(EC.presence_of_element_located((By.ID, "_com_liferay_login_web_portlet_LoginPortlet_account")))
    user_input.clear()
    user_input.send_keys(username)
    print("⌨️ Digitando el Username")
        
    pass_input = wait.until(EC.presence_of_element_located((By.ID, "_com_liferay_login_web_portlet_LoginPortlet_password")))
    pass_input.clear()
    pass_input.send_keys(password)
    print("⌨️ Digitando el Password")

    # btn = wait.until(EC.element_to_be_clickable((By.ID, "_com_liferay_login_web_portlet_LoginPortlet_nlfq")))
    # driver.execute_script("arguments[0].click();", btn)

    boton = driver.find_element(By.XPATH, "//button[span[text()='Acceder']]")
    boton.click()
    print("🖱️ Clic en 'Acceder'.")

    time.sleep(3)

    elemento = wait.until(EC.element_to_be_clickable((By.ID, "bg")))
    #elemento.click()
    driver.execute_script("arguments[0].click();", elemento)
    print("🖱️ Clic en 'Menu Lateral'.")

    ele = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[span[contains(text(), 'Consulta de pólizas')]]")))
    driver.execute_script("arguments[0].click();", ele)
    print("🖱️ Clic en 'Consulta de polizas'.")

    while True:

        try:
            todos_los_datos = consultarAPI(url_datos_para_cancelar_cuotas,ids_compania)

            if not todos_los_datos:
                raise Exception("No se recibió información de ninguna compañía.")

            log_path,ruta_salida_API,ruta_salida,ruta_maestro,nombre_hoja, ruta_carpeta_facturas, ruta_carpeta_comprobante, ruta_carpeta_errores,carpeta_compañia,carpeta_principal = crearCarpetas(todos_los_datos,nombre_carpeta_compañia,tipo=2,cia_a_verificar=None)

            print("\n📁 Iniciando procesamiento para Mapfre...")

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
