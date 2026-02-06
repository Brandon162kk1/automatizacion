#-- Imports ---
import subprocess
import time
import os
import pandas as pd
import requests
import socket
import shutil
#-- Froms ----
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver import ActionChains
#from excels.estilosExcel import guardar_excel_con_formato
from Sunat.validar_factura import consultarValidezSunat,login_sunat
from Birlik.cancelar_cuotas import agregar_comprobante_pago,cancelar_y_agregar_cuota,url_cuotas_canceladas,url_datos_para_cancelar_cuotas
from Apis.api_birlik import consultarAPI
from GoogleChrome.chromeDriver import abrirDriver, crearCarpetas
from GoogleChrome.fecha_y_hora import get_timestamp
from datetime import datetime
from selenium.webdriver.support.ui import WebDriverWait

#--------- COMPAÑÍA SANITAS MAPFRE ------
ruc_mapfre_salud = '20517182673' # Salud
ruc_mapfre_pension = '20418896915' #Seguros y Re aseguros
# Lista de IDs de compañía
ids_compania = [16,17,18]
#-------CREDENCIALES SANITAS----------
login_url_mapfre = os.getenv("url_mapfre")
username = os.getenv("usernameMapfre")
password = os.getenv("passwordMapfre")
nombre = os.getenv("CONT_NAME", socket.gethostname())
nombre = os.getenv("CONT_NAME")
#--- NOMBRE SERVICE DNS DOCKER PARA UTILIZAR EN LA API -----
nom_serv = os.getenv("nom_serv")
#----- Carpeta de la Compañia -------
nombre_carpeta_compañia = f"Mapfre_{get_timestamp()}"

def limpiar(valor):
    if valor is None:
        return ""
    valor = valor.strip()
    return valor if valor else ""

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
        print("❌ Error durante el flujo de descarga:", ex)
        driver.save_screenshot(f"{ruta_carpeta_errores}/{numero_poliza}_ventanalinux.png")
        return False

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

    if fk_compania_birlik == '17':
        ruc_compania = ruc_mapfre_pension
    else:
        ruc_compania = ruc_mapfre_salud

    resultado_importe = False
    resultado_sunat = False
    resultado_birlik = False
    resultado_estado = None
    resultado_accion = ""

    # if id_cuota_birlik not in ("10802"):
    #      return f"Pagina Web en Mantenimiento" if resultado_importe else "Pagina Web en Mantenimiento" ,"Pagina Web en Mantenimiento" if resultado_sunat else "Pagina Web en Mantenimiento" ,"Pagina Web en Mantenimiento" if resultado_birlik else "Pagina Web en Mantenimiento" , "" if resultado_estado else "Pagina Web en Mantenimiento", "" if resultado_accion else "Pagina Web en Mantenimiento"

    try:
        # ------------------ Inicio del Flujo de Automatización ------------------
        label = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-label[normalize-space()='Nro. Póliza *']")))
        driver.execute_script("arguments[0].click();", label)
        #time.sleep(2)
        print("🖱️ Clic en el # de Póliza")

        poliza_input = wait.until(EC.presence_of_element_located((By.ID, "mat-input-5")))
        poliza_input.clear()
        poliza_input.send_keys(numero_poliza_birlik)
        print(f"⌨️ Digitando la póliza '{numero_poliza_birlik}'.")

        time.sleep(3)

        buscar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Buscar']")))
        buscar_btn.click()
        print("🖱️ Clic en 'Buscar'.")

        try:
            print("⌛ Esperando que cargue la tabla...")
            wait.until(
                EC.presence_of_element_located((
                    By.XPATH, "(//ul[contains(@class,'g-tbl-row')])[2]"
                ))
            )
            print("📄 La tabla cargó correctamente.")
        except:
            raise Exception("No cargó la tabla")

        encontrado = False

        while True:

            filas = driver.find_elements(By.XPATH, "//ul[contains(@class,'g-tbl-row')]")
            filas = filas[1:]

            for fila in filas:

                columnas = fila.find_elements(By.TAG_NAME, "li")

                # Asegurar que la fila tenga suficientes columnas
                if len(columnas) < 4:
                    continue

                # CodigoCuota
                codigo_fila = limpiar(columnas[6].text)
                importe_fila = limpiar(columnas[7].text)
                fecha_emision_fila = limpiar(columnas[9].text)
                numero_factura = limpiar(columnas[11].text)
                fecha_pago_fila = limpiar(columnas[14].text)

                if codigo_fila == numero_proforma_birlik:

                    encontrado = True
                    print(f"✅ Fila encontrada ")
                    placeholder = "-"
                    print(
                        f"⌛ Código Cuota: {codigo_fila or placeholder}, "
                        f"Importe: {importe_fila or placeholder}, "
                        f"Fecha Emisión: {fecha_emision_fila or placeholder}, "
                        f"Numero Factura: {numero_factura or placeholder}, "
                        f"Fecha de Pago: {fecha_pago_fila or placeholder}"
                    )
                    diferencia = abs(float(importe_fila) - float(importe_total_birlik))
                    print(f"Importe de Birlik: {float(importe_total_birlik)} -- Importe de la Compañía : {float(importe_fila)}")
                    if diferencia > 0.05:
                        print("❌ Los importes No coinciden")
                    else:
                        print("✅ Los importes Coinciden")
                        resultado_importe = True
                    
                    if numero_factura == '-' or fecha_emision_fila == '-' or fecha_pago_fila == '-':
                        resultado_estado = "No dice estado"
                        resultado_accion = "Esperar"
                        break

                    # OPCIONAL: marcar checkbox
                    checkbox = fila.find_element(By.XPATH, ".//input[@type='checkbox']")
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", checkbox)
                    driver.execute_script("arguments[0].click();", checkbox)
                    print("🖱️ Clic en el 'Checkbox'.")

                    time.sleep(2)

                    ventana_principal_mapfre = driver.current_window_handle

                    ruta_factura = os.path.join(ruta_carpeta_facturas, f"{numero_poliza_birlik}_{numero_factura}.pdf")
                    destino_factura = f"{ruta_carpeta_facturas}/{numero_poliza_birlik}_{numero_factura}"

                    for i in range(3):   # porque aparecen 3 archivos
    
                        if i == 0:
                            print("📥 Descargando Factura (1/3)...")

                            if click_descarga_factura(driver, destino_factura, columnas[11], numero_poliza_birlik, ruta_carpeta_errores):
                                print("✅ Factura descargada")
                            else:
                                raise Exception("❌ Error en descarga de factura principal")

                            time.sleep(2)

                        else:
                            print(f"⌛ Cancelando descarga automática {i+1}/3...")
        
                            # Cancelar Save File Dialog
                            subprocess.run(["xdotool", "key", "Escape"])
                            print("⌨️ Se presionó ESC para cancelar la descarga.")

                            time.sleep(1)
                
                    # Lista para guardar las fechas
                    fecha_habiles_factura = []

                    fecha_emision_probar = datetime.strptime(fecha_pago_fila, "%d/%m/%Y")

                    fecha_habiles_factura.append(fecha_emision_probar.strftime("%d/%m/%Y"))

                    print(f"📅 Fechas de Emisión a probar:{fecha_habiles_factura}")

                    for fecha in fecha_habiles_factura:
                        print("---------------------------------------")
                        print(f"⌛ Probando con la Fecha hábil: {fecha}")

                        #------------INGRESA A SUNAT-------  
                        nombre_imagen_sunat = f"{numero_proforma_birlik}_{numero_poliza_birlik}.png"
                        ruta_imagen_sunat = os.path.join(ruta_carpeta_comprobante, nombre_imagen_sunat)
                        resultado = consultarValidezSunat(driver,wait,ruc_compania,tipo_doc_birlik,ruc_cliente_birlik,numero_factura,fecha,importe_fila,ruta_imagen_sunat)

                        driver.switch_to.window(ventana_principal_mapfre)
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
                                # Subir comprobante a Birlik
                                print("📤 Subiendo todos los documentos a Birlik...")
                                cancelar_y_agregar_cuota(driver,wait,id_cuota_birlik,numero_factura,fecha,ruta_factura,ruta_imagen_sunat,resultado_importe)
                                resultado_accion = f'=HYPERLINK("{url_cuotas_canceladas}{fk_Cliente_birlik}", "Enviar Factura")'
            
                            resultado_birlik = True
                            break  # Salir del bucle porque ya funcionó con esa fecha

                        else:
                            resultado_accion = f'=HYPERLINK("{login_sunat}", "Ver Sunat")'
                            continue # Si no es True, salta al siguiente intento

                    break              

            if encontrado:
                break

            if not encontrado:
                try:
                    # Buscar el botón sin lanzar error
                    btn_siguiente_list = driver.find_elements(By.XPATH, "//button[@aria-label='Siguiente página']")

                    # Si no existe el botón
                    if not btn_siguiente_list:
                        raise Exception(f"No existe el botón 'Siguiente página' porque no hay resultados para la póliza '{numero_poliza_birlik}'")

                    # Tomar el botón encontrado
                    btn_siguiente = btn_siguiente_list[0]

                    # Si está deshabilitado → NO HAY MÁS PÁGINAS
                    if btn_siguiente.get_attribute("disabled"):
                        raise Exception("No hay más páginas")

                    # Hacer scroll y clic
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});",btn_siguiente)
                    time.sleep(2)
                    btn_siguiente.click()

                    print("🔄 Cambiando a la siguiente página...")
                    time.sleep(2)

                except Exception as e:
                    print(f"⛔ {e}")
                    break

        if not encontrado:
            resultado_estado = "No se encuentro cuota"
            resultado_accion = f'=HYPERLINK("{url_cuotas_canceladas}", "Revisar Cuota")'
            raise Exception(f"No se encontró ninguna fila con el código {numero_proforma_birlik}")

    except Exception as e:
        print(f"❌ Error Procesando toda la fila, Motivo: {e}")
    finally:
        limpiar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Limpiar']")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", limpiar_btn)
        time.sleep(0.3)
        wait.until(lambda d: limpiar_btn.location['y'] > 0)
        limpiar_btn.click()
        print("🖱️ Clic en 'Limpiar'.")
        return f"Coinciden" if resultado_importe else f"No coinciden" ,"Válido" if resultado_sunat else "No Válido" ,"Cuota Cancelada" if resultado_birlik else "Cuota Pendiente" , "No indica" if resultado_estado is None else resultado_estado ,resultado_accion

def main():
    
    display_num = os.getenv("DISPLAY_NUM", "0")
    os.environ["DISPLAY"] = f":{display_num}"

    driver, wait = abrirDriver()

    # ------------------ Inicio del Flujo de Automatización ------------------
    driver.get(login_url_mapfre) 
    print("⌛ Ingresando a la URL")

    user_input = wait.until(EC.presence_of_element_located((By.ID, "mat-input-1")))
    user_input.clear()
    user_input.send_keys(username)
    print("⌨️ Digitando el Username")
        
    pass_input = wait.until(EC.presence_of_element_located((By.ID, "mat-input-0")))
    pass_input.clear()
    pass_input.send_keys(password)
    print("⌨️ Digitando el Password")

    ingresar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Ingresar')]")))
    ingresar_btn.click()
    print("🖱️ Clic en 'Ingresar'.")

    # # Esperar hasta que el botón Brokers sea clickeable
    # brokers_btn = wait.until(EC.element_to_be_clickable(
    #     (By.XPATH, "//a[.//span[contains(text(),'Brokers')]]")
    # ))
    # # Hacer clic
    # brokers_btn.click()
    # print("🖱️ Se hizo clic en 'Brokers'")

    elemento = wait.until(EC.element_to_be_clickable(
        (By.CSS_SELECTOR, ".card-modality__item--left")
    ))
    elemento.click()
    print("🖱️ Clic en enviar por Correo Electronico.")

    codigo_mapfre_path = "/codigo_mapfre/codigo.txt"

    while not os.path.exists(codigo_mapfre_path):
        time.sleep(2)

    with open(codigo_mapfre_path, "r") as f:
        codigo = f.read().strip()

    print(f"✅ Código recibido desde volumen: {codigo}")

    inputs = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input.g-input-codes__code")))

    if len(inputs) == len(codigo):
        for i, inp in enumerate(inputs):
            inp.clear()
            inp.send_keys(codigo[i])
    else:
        raise Exception("Los inputs no coinciden con la longitud del código")

    time.sleep(1)

    try:
        comprobar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Comprobar')]")))
        comprobar_btn.click()
        print("🖱️ Clic en 'Comprobar'")
    except Exception as e:
        raise Exception(f"No se pudo hacer clic en Comprobar -> {e}")

    # --- Eliminar el archivo después de usarlo ---
    try:
        os.remove(codigo_mapfre_path)
        print("🧹 Archivo codigo.txt eliminado desde volumen")
    except FileNotFoundError:
        print("⚠️ No se encontró codigo.txt al intentar eliminarlo (ya fue borrado)")
    except Exception as e:
        print(f"❌ Error al eliminar codigo.txt: {e}")

    try:
        boton_ok = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(), 'Ok')]]")))
        boton_ok.click()
        print("🖱️ Clic en 'Ok'")
    except TimeoutException:
        pass

    consulta_gestion = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='CONSULTAS DE GESTION']")))
    consulta_gestion.click()
    print("🖱️ Clic en 'CONSULTAS DE GESTION'.")
          
    action = ActionChains(driver)   
    cobranzas_link = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[normalize-space()='COBRANZAS']")))
    action.move_to_element(cobranzas_link).perform()
    print("🖱️ Mouse sobre 'COBRANZAS'.")
  
    cronograma_pagos = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='CRONOGRAMA DE PAGOS']")))
    action.move_to_element(cronograma_pagos).click().perform()
    print("🖱️ Clic en 'CRONOGRAMA DE PAGOS'.")

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
