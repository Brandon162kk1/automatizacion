#-- Imports ---
import subprocess
import time
import os
import pandas as pd
import pdfplumber
import glob
import re
import shutil
#-- Froms ----
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime, timedelta
#from excels.estilosExcel import guardar_excel_con_formato
from Sunat.validar_factura import consultarValidezSunat,login_sunat
from Birlik.cancelar_cuotas import agregar_comprobante_pago,cancelar_y_agregar_cuota,url_cuotas_canceladas,url_detalle_poliza,url_datos_para_cancelar_cuotas
from Apis.api_birlik import consultarAPI
from GoogleChrome.chromeDriver import abrirDriver, crearCarpetas
from GoogleChrome.fecha_y_hora import get_timestamp,get_fecha_hoy
# from tkinter.tix import CELL

#--------- COMPAÑÍA PACIFICO ------
ids_compania = [23,33,24]            #-- > 24 es SALUD , 33 es Vida Ley , 23 es PACIFICO GENERAL (PENSION)
ruc_pacifico_vida = '20332970411'   #--> Seguros de Vida
ruc_pacifico_salud = '20431115825'  #--> Salud EPS
#----- Carpeta de la Compañia -------
nombre_carpeta_compañia = f"Pacifico_{get_timestamp()}"
#----- Variables de Entorno -------
urlPacifico = os.getenv("url_pacifico")
correo = os.getenv("remitente")
password = os.getenv("passwordCorreo")

# Función para limpiar el valor y convertirlo a float o int
def limpiar_valor(valor):
    # Eliminar símbolos de moneda, comas, espacios, etc.
    limpio = valor.replace("S/", "").replace("US$", "").replace(",", "").replace(" ","").strip()
    # Convertir a float o int según convenga
    return float(limpio) if "." in limpio else int(limpio)

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

    if fk_compania_birlik == '24':
        ruc_emisor = ruc_pacifico_salud
    elif fk_compania_birlik == '23' and fk_Ramo_birlik == '55': #SI O SI TODO PENSION ESTA CON ESTE RUC
        ruc_emisor = ruc_pacifico_vida
    else:
        ruc_emisor = ruc_pacifico_vida

    resultado_importe = False
    resultado_sunat = False
    resultado_birlik = False
    resultado_estado = None
    resultado_accion = ""
    cod_giro_comparar = None

    # if numero_proforma_birlik != "1228638760":
    #     return f"Pagina Web en Mantenimiento" if resultado_importe else "Pagina Web en Mantenimiento" ,"Pagina Web en Mantenimiento" if resultado_sunat else "Pagina Web en Mantenimiento" ,"Pagina Web en Mantenimiento" if resultado_birlik else "Pagina Web en Mantenimiento" , "" if resultado_estado else "Pagina Web en Mantenimiento", "" if resultado_accion else "Pagina Web en Mantenimiento"

    try:

       # ------------------ Inicio del Flujo de Automatización ------------------ 
       for i in range(2):
           try:
                btn_aceptar = WebDriverWait(driver,1.5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.pga-alert-close")))
                btn_aceptar.click()
                time.sleep(3)
                print("🖱️ Se detectó modal y se hizo clic en 'Aceptar'.")
                break
           except TimeoutException:
                pass
        
       cod_giro_comparar = []
       fecha_habiles_factura = []

       if fk_compania_birlik == '24': #SALUD

          cod_giro_comparar,fecha_emision_factura = obtener_cod_giro_SCTR(driver,wait,numero_poliza_birlik,
                                                                          tipo_doc_birlik,ruc_cliente_birlik,
                                                                          id_cuota_birlik,fk_Cliente_birlik,
                                                                          numero_proforma_birlik,importe_total_birlik,
                                                                          estadoCuota_birlik,primaneta_birlik,ruta_carpeta_facturas,
                                                                          ruta_carpeta_errores,fecha_inicioVig_Birlik,fecha_finVig_Birlik)
          fecha_emision_factura2 = datetime.strptime(fecha_emision_factura, "%d/%m/%Y")
          fecha_habiles_factura.append(fecha_emision_factura2.strftime("%d/%m/%Y"))

       else: # PENSION y VIDA LEY , EL CODIGO DE CUOTA LO GUARDO EN UNA LISTA
          cod_giro_comparar.append(numero_proforma_birlik)

       if not cod_giro_comparar:
          raise Exception("No se pudo obtener los códigos de SCTR.")

       menu_estados = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.dropdown-toggle.EstadosdeCuenta")))
       menu_estados.click()
       print("🖱️ Clic en Estado de cuenta")

       link_cuotas = wait.until(EC.element_to_be_clickable((By.XPATH, "//ul[contains(@class,'dropdown-menu')]//a[b[text()='Cuotas']]")))
       link_cuotas.click()
       print("🖱️ Clic en 'Cuotas'.")

       poliza_input = wait.until(EC.presence_of_element_located((By.ID, "inputBuscador")))
       poliza_input.clear()
       poliza_input.send_keys(numero_poliza_birlik)
       print(f"✅ Póliza ingresada: {numero_poliza_birlik}")

       buscar_btn = wait.until(EC.element_to_be_clickable((By.ID, "busqueda")))
       buscar_btn.click()
       print("🖱️ Clic en 'Buscar'.")
  
       fila_encontrada = False 
       fila_menos_columna = False
       fecha_antigua_bus = False

       while True:

           print("⌛ Esperando que cargue la tabla...")

           wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table[@id='tablaCuota']//tbody/tr")))
           rows = driver.find_elements(By.XPATH, "//table[@id='tablaCuota']//tbody/tr")
           print(f"🔍 Filas detectadas: {len(rows)}")

           fila_encontrada_codCuota = False

           for i, row in enumerate(rows):

               cells = row.find_elements(By.TAG_NAME, "td")

               if len(cells) < 9:
                   print(f"⚠️ Fila {i} ignorada (tiene {len(cells)} columnas), Posiblemente no tiene registros.")
                   resultado_estado = "No está registrado aun"
                   resultado_accion = f'=HYPERLINK("{url_cuotas_canceladas}", "Revisar Cuota")'
                   # driver.save_screenshot(f"{ruta_carpeta_errores}/fila_ignorada_{numero_poliza_birlik}_{fk_Cliente_birlik}.png")
                   continue
                
               # codigoCuota o Liquidacion de la prima → columna 7 (índice 6)
               aviso_cobranza = cells[6].text.strip()      # para salud aviso de cobranza
               num_cuota_liquida = cells[7].text.strip()   # para salud
               importe_dolares= cells[8].text.strip()
               importe_soles = cells[9].text.strip()
               #vig_ini_birlik = cells[4].text.strip()

               # # Logica para convertir la fecha de vigencia de la poliza en la CIA a formato datetime
               # dia0, mes0, anio0 = vig_ini_birlik.split("/")

               # if len(anio0) == 2:
               #     anio0 = "20" + anio0  # agrega el '20' delante

               # fecha_completa0 = f"{dia0}/{mes0}/{anio0}"

               # fecha_emision_probar0 = datetime.strptime(fecha_completa0, "%d/%m/%Y")
               # #-------------------------------------------------------------------------------------

               # PARA VIDA LEY , EL AVISO DE COBRANZA Y LQUIDACION SON EL MISMO VALOR
               # PARA SALUD SON DIFERENTES EL AVISO DE COBRANZA Y LIQUIDACIOMN , CUALQUIERA DE AMBOS PUEDE SER EL CODIGO DE CUOTA
               # PARA PENSION , EL NUMERO DE LIQUIDACION ES EL CODIGO DE CUOTA CONSIDERANDO QUE PUEDE SER 1 DE 6
               print(f"⌛ Avs.Cobranza -> {aviso_cobranza}, NumCuota Liqui -> {num_cuota_liquida}, Imp Dolares $ -> {importe_dolares}, Imp Soles S/. -> {importe_soles}")

               if fk_compania_birlik == '33':
                  condicion = aviso_cobranza in cod_giro_comparar
               elif fk_compania_birlik == '24':
                  condicion = num_cuota_liquida in cod_giro_comparar #--> STRING DENTRO DE UNA LISTA
               else:
                  num_cuota_liquida = re.sub(r"\(.*\)", "", num_cuota_liquida).strip()
                  condicion =  num_cuota_liquida in cod_giro_comparar             

               if condicion:

                   fila_encontrada_codCuota = True
                   fila_encontrada = True

                   # 1. Clic en la celda para expandir
                   primera_celda = row.find_elements(By.TAG_NAME, "td")[0]
                   time.sleep(2)
                   driver.execute_script("arguments[0].click();", primera_celda)
                   print(f"🖱️ Clic en la primera celda para la Fila {i}")

                   time.sleep(2)

                   child_row = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "tr.child")))
                   print("✅ Fila hija cargada:")
                   time.sleep(2)
                   estado = child_row.find_element(By.XPATH, ".//span[@class='dtr-title' and text()='Estado']/following-sibling::span").text.strip()
                   resultado_estado = estado #-> Asignando el Estado para que se muestre en el Excel
                   links = child_row.find_elements(By.CSS_SELECTOR, "a.donwloadComprobante")
                   fecha_pago = child_row.find_element(By.XPATH, ".//span[@class='dtr-title' and contains(text(),'F. Pago')]/following-sibling::span").text.strip()

                   importe_raw = importe_soles if importe_dolares == '-' else importe_dolares
                   importe = re.sub(r"[^\d.,]", "", importe_raw)

                   if importe == '-':
                       raise Exception ("El importe ya no se muestra.")

                   print("----------------------------------------")
                   print(f"✅ Fila encontrada: Póliza ->'{numero_poliza_birlik}', Código Cuota ->'{num_cuota_liquida}', Importe -> '{importe}'.")

                   diferencia = abs(float(importe) - float(importe_total_birlik))
                   print(f"Importe de Birlik: {float(importe_total_birlik)} -- Importe de la Compañía : {float(importe)}")
                   if diferencia > 0.05:
                       print("❌ Los importes No coinciden")
                   else:
                       print("✅ Los importes Coinciden")
                       resultado_importe = True

                   if links:

                       link_comprobante = links[0]
                       factura = link_comprobante.text.strip()   # F028-0001018980
                       prefijo, numero = factura.split("-")      # → "F028", "0001018980"
                       numero_mod = numero[2:]                   # quitar los 2 primeros caracteres
                       factura_final = f"{prefijo}-{numero_mod}"

                   else:
                       factura_final = "-"                       # no hay link, solo hay texto '-'
                       if estado.lower() == "pagada":
                           resultado_accion = f'=HYPERLINK("{url_cuotas_canceladas}", "Revisar Cuota")'
                           raise Exception("No se encontró el elemento de descarga (a.donwloadComprobante), aunque el estado es 'pagada'. No se puede proceder con la descarga.")

                   print(f"🧾 Factura: {factura_final}")
                   print(f"📅 Fecha Pago: {fecha_pago}")

                   if estado.lower() == "pagada":
                       
                       print(f"🟢 Estado: {estado}")

                       ventana_principal_pacifico = driver.current_window_handle

                       link_comprobante = child_row.find_element(By.CSS_SELECTOR, "a.donwloadComprobante")

                       #---- Ruta para eliminar la pre-factura apenas procese otra fila
                       ruta_factura = os.path.join(ruta_carpeta_facturas, f"{numero_poliza_birlik}_{factura_final}.pdf")
                       destino_factura = f"{ruta_carpeta_facturas}/{numero_poliza_birlik}_{factura_final}"

                       if click_descarga_factura(driver, destino_factura,link_comprobante,numero_poliza_birlik,ruta_carpeta_errores):
                           print("📥 Descarga correcta")
                       else:
                           raise Exception("Error en descarga de factura")

                       try:

                           dia, mes, anio = fecha_pago.split("/")

                           if len(anio) == 2:
                               anio = "20" + anio  # agrega el '20' delante

                           fecha_completa = f"{dia}/{mes}/{anio}"

                           fecha_emision_probar = datetime.strptime(fecha_completa, "%d/%m/%Y")

                           for i in range(-3, 4):  # Desde -3 hasta +3 (incluye 0 → hoy)
                               fecha = fecha_emision_probar + timedelta(days=i)

                               # Si la siguiente fecha es mayor que hoy, se detiene
                               if fecha.date() >= get_fecha_hoy().date():
                                   break

                               fecha_habiles_factura.append(fecha.strftime("%d/%m/%Y"))
                        
                           for fecha in fecha_habiles_factura: #(*)

                                   print("---------------------------------------")
                                   print(f"⌛ Probando con la Fecha hábil: {fecha}")

                                   nombre_imagen_sunat = f"{numero_proforma_birlik}_{numero_poliza_birlik}.png"
                                   ruta_imagen_sunat = os.path.join(ruta_carpeta_comprobante, nombre_imagen_sunat)
                                   resultado = consultarValidezSunat(driver,wait,ruc_emisor,tipo_doc_birlik,ruc_cliente_birlik,factura_final,fecha,importe_total_birlik,ruta_imagen_sunat)

                                   driver.switch_to.window(ventana_principal_pacifico)
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
                                            cancelar_y_agregar_cuota(driver,wait,id_cuota_birlik,factura_final,fecha,ruta_factura,ruta_imagen_sunat,resultado_importe)
                                            resultado_accion = f'=HYPERLINK("{url_cuotas_canceladas}{fk_Cliente_birlik}", "Enviar Factura")'
         
                                       resultado_birlik = True

                                       break   #(*)

                                   else:
                                       resultado_accion = f'=HYPERLINK("{login_sunat}", "Ver Sunat")'
                                       continue

                       finally:
                           os.remove(ruta_factura)

                   elif estado.lower() == "por vencer":
                       print(f"🟡 Estado : {estado}")
                       resultado_accion = f'Sin Observación'
                   else:
                       print(f"🛑 Estado : {estado}")
                       resultado_accion = f'=HYPERLINK("{url_detalle_poliza}{id_Poliza_birlik}", "Anular Cuota")'

                   break

           # Si la fila se encontró → salir del while
           if fila_encontrada_codCuota:
               break

           # if fila_menos_columna:
           #     break

           # if fecha_antigua_bus:
           #     break

           # --- Si NO se encontró fila → intentar pasar de página ---
           print("➡ Pasando a la siguiente página de la Tabla")
                        
           try:
               if len(rows) < 10:
                   driver.save_screenshot(f"{ruta_carpeta_errores}/noEncontrofila_{numero_poliza_birlik}_{numero_proforma_birlik}_{get_timestamp()}.png")
                   raise Exception("No hay más páginas")
               else:
                   btn_next_li = driver.find_element(By.ID, "tablaCuota_next")

                   if "disabled" in btn_next_li.get_attribute("class"):
                       raise Exception("Botón SIGUIENTE deshabilitado. Ya no hay más páginas.")

                   btn_next_a = btn_next_li.find_element(By.TAG_NAME, "a")
                   driver.execute_script("arguments[0].click();", btn_next_a)

                   print("➡ Clic en botón SIGUIENTE")
           except Exception as e:
               print(f"⛔ {e}")
               break

           # esperar a que la tabla recargue
           time.sleep(4)

       if not fila_encontrada:
           raise Exception (f"No se encontró la cuota {numero_proforma_birlik} en ninguna página.")

    except Exception as e:
        print(f"❌ Error Procesando toda la fila, Motivo: {e}")
    finally:
        menu_principal = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='Pólizas']")))
        menu_principal.click()
        return f"Coinciden" if resultado_importe else f"No coinciden" ,"Válido" if resultado_sunat else "No Válido" ,"Cuota Cancelada" if resultado_birlik else "Cuota Pendiente" ,resultado_estado,resultado_accion

def obtener_cod_giro_SCTR(driver,wait,numero_poliza_birlik,tipo_doc_birlik,ruc_cliente_birlik,id_cuota_birlik,fk_Cliente_birlik
                                              ,numero_proforma_birlik,importe_total_birlik,estadoCuota_birlik,
                                              primaneta_birlik,ruta_carpeta_facturas,ruta_carpeta_errores,fecha_inicioVig_Birlik,
                                              fecha_finVig_Birlik):

    codigo_giro = []
    fechas_factura = None

    try:

            poliza_input = wait.until(EC.presence_of_element_located((By.ID, "inputBuscador")))
            poliza_input.clear()
            poliza_input.send_keys(numero_poliza_birlik)
            print(f"✅ Póliza ingresada: {numero_poliza_birlik}")

            buscar_btn = wait.until(EC.element_to_be_clickable((By.ID, "busqueda")))
            buscar_btn.click()
            print("🖱️ Clic en 'Buscar'.")

            wait.until(EC.visibility_of_element_located((By.ID,"tablaPoliza")))
            print("⌛ Esperando que cargue la tabla...")

            try:
                wait.until(lambda d: len(d.find_elements(By.XPATH, "//table[@id='tablaPoliza']//tr")) > 1)
                print("✅ La tabla tiene al menos 1 fila.")
            except TimeoutException:
                print("⏰ No se cargaron las filas en el tiempo esperado.")

            table = driver.find_element(By.ID, "tablaPoliza")
            rows = table.find_elements(By.XPATH, ".//tbody//tr")
        
            fila_encontrada = False
   
            for i, row in enumerate(rows): #(***)

                cells = row.find_elements(By.TAG_NAME, "td")

                if len(cells) < 11:
                    print(f"⚠️ Fila {i} ignorada (tiene {len(cells)} columnas)")
                    continue
     
                poliza_cia = cells[3].text.strip()
                inicio_vigencia_cia = cells[5].text.strip()
                fin_vigencia_cia = cells[6].text.strip()
                estado_cia = cells[9].text.strip()
               
                if numero_poliza_birlik == poliza_cia:

                    fila_encontrada = True

                    print(f"✅ Fila encontrada: Póliza='{poliza_cia}', Estado='{estado_cia}', Inicio de Vigencia='{inicio_vigencia_cia}', Fin de Vigencia='{fin_vigencia_cia}'.")
                
                    enlace_poliza = cells[3].find_element(By.TAG_NAME, "a")
                    wait.until(EC.element_to_be_clickable(enlace_poliza))
                    driver.execute_script("arguments[0].click();", enlace_poliza)
                    print("🖱️ Clic (por JS) en la póliza")

                    print("⌛ Cargando...")
                    try:
                        wait.until(EC.visibility_of_element_located((By.ID, "ajax-loading")))
                    except TimeoutException:
                        pass
                    finally:
                        print("✅ Loader desapareció, continuando...")

                    tab_gestion = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#idTabGestion a")))
                    driver.execute_script("arguments[0].click();", tab_gestion)
                    print("🖱️ Clic en Gestión de Póliza")

                    try:
                        wait.until(EC.invisibility_of_element_located((By.XPATH, "//div[@class='titulo' and contains(., 'Un momento por favor')]")))
                        print("⌛ Cargando")
                    except:
                        print("✅ Ya desapareció antes de que empiece a esperar")

                    span = "¿Deseas realizar alguna consulta o requerimiento?"
                    link = wait.until(EC.element_to_be_clickable((By.XPATH, f"//a[normalize-space(text())='{span}']")))
                    link.click()
                    print(f"🖱️ Clic en '{span}'.")

                    btn_buscar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space(text())='Buscar contrato/póliza']")))
                    btn_buscar.click()
                    print(f"🖱️ Clic en el botón Buscar contrato/Póliza. ")

                    ruc_input = wait.until(EC.presence_of_element_located((By.ID, "search")))
                    ruc_input.clear()
                    ruc_input.send_keys(ruc_cliente_birlik)
                    print(f"✅ Ruc ingresado: {ruc_cliente_birlik}")

                    dropdown_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.niche-dropdown-button")))
                    dropdown_btn.click()

                    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.niche-dropdown-content")))

                    opcion_all = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@value='All']")))
                    driver.execute_script("arguments[0].click();", opcion_all)  # <-- evita overlays
                    print("🖱️ Clic en 'Todos'.")

                    time.sleep(3)

                    btn_buscar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.medium")))
                    btn_buscar.click()
                    print(f"🖱️ Clic en Buscar")

                    poliza2_input = wait.until(EC.presence_of_element_located((By.ID, "secondSearch")))
                    poliza2_input.clear()
                    poliza2_input.send_keys(numero_poliza_birlik)
                    print(f"✅ Póliza nuevamente ingresada: {numero_poliza_birlik}")

                    btn_buscar_2 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.small")))
                    btn_buscar_2.click()
                    print(f"🖱️ Clic en Buscar")
               
                    btn_poliza = wait.until(EC.element_to_be_clickable((By.XPATH, f"//button[@class='button-continue' and normalize-space(text())='{numero_poliza_birlik}']")))
                    btn_poliza.click()
                    print(f"🖱️ Clic en la póliza {numero_poliza_birlik}.")

                    enlace_facturtacion = wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[@class='enlace' and contains(., 'Facturación')]")))
                    enlace_facturtacion.click()
                    print(f"🖱️ Clic en 'Facturación'.")
                    time.sleep(3)

                    fila_encontrada_codCuota = False

                    while True: #*(While)

                        tabla = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".sctr-pagination.center")))

                        rows = tabla.find_elements(By.CSS_SELECTOR, ".sctr-table-body-pagination")

                        for i in range(len(rows)): #(**)

                            ventana_principal_pacifico = driver.current_window_handle

                            time.sleep(3)

                            # 🔄 REFETCH: evita stale y evita descargar repetidos
                            rows = tabla.find_elements(By.CSS_SELECTOR, ".sctr-table-body-pagination")
                            row = rows[i]

                            cols = row.find_elements(By.CSS_SELECTOR, ":scope .column-regular")

                            fecha_emision = cols[0].text.strip() #esta fecha es de la factura
                            numero_documento = cols[4].text.strip()
                            estado_cuota = cols[7].text.strip()
                            monto = cols[2].text.strip() # Para Salud esta columna es su Prima Neta
                            monto_prima_neta = float(monto.replace(',', ''))

                            #---- Ruta para eliminar la pre-factura apenas procese otra fila
                            ruta_factura = os.path.join(ruta_carpeta_facturas, f"{numero_poliza_birlik}_{numero_documento}.pdf")
                        
                            print(f"⌛ Fila {i+2}: Prima Neta -> {monto_prima_neta}, Fecha Emisión -> {fecha_emision}, Factura -> {numero_documento}, Estado Cuota -> {estado_cuota}")

                            if datetime.strptime(fecha_emision, "%d/%m/%Y") < datetime.strptime(fecha_inicioVig_Birlik, "%d/%m/%Y"):
                                raise("La Fecha de Emisión para esa Cuota ya es demasiado antigua.")

                            if float(monto_prima_neta) == float(primaneta_birlik):

                                print("--------------------------------------------------")

                                print(f"🔎 Probando fila con la Factura: {numero_documento}")
                      
                                # Clic en Ver documento (columna 8) si es Salud
                                boton_detalle = row.find_element(By.CSS_SELECTOR, ".btn-border")
                                driver.execute_script("arguments[0].click();", boton_detalle)

                                time.sleep(3)

                                iframe = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "iframe[src^='data:application/pdf']")))

                                driver.switch_to.frame(iframe)

                                open_btn = wait.until(EC.element_to_be_clickable((By.ID, "open-button")))

                                destino_factura = f"{ruta_carpeta_facturas}/{numero_poliza_birlik}_{numero_documento}"

                                try:
                                    if click_descarga_factura(driver, destino_factura, open_btn,numero_poliza_birlik,ruta_carpeta_errores):
                                        print("📥 Descarga correcta")
                                    else:
                                        raise Exception("Error en descarga de factura")
                                finally:

                                    time.sleep(2)

                                    # Salir del iframe y volver al contenido principal
                                    driver.switch_to.default_content()

                                    try:
                               
                                        actions = ActionChains(driver)
                                        actions.move_by_offset(10, 10).click().perform()
                                        actions.move_by_offset(-10, -10).perform()
                                        #print("Modal cerrado.1")
                                    except :

                                        try:
                                            backdrop = driver.find_element(By.CSS_SELECTOR, "div[style*='position: fixed']")
                                            driver.execute_script("""
                                                arguments[0].style.display='none';
                                            """, backdrop)
                                            print("Modal cerrado.")
                                        except :
                                            driver.execute_script("""
                                            document.querySelectorAll('.modal, .modal-overlay, .MuiDialog-root, .MuiBackdrop-root')
                                              .forEach(e => e.remove());
                                            """)
                                            print("Modal cerrado.2")

                                    time.sleep(2)

                                    driver.switch_to.window(ventana_principal_pacifico)
                                    print("🔄 Volviendo a la pestaña principal --")

                                id_prof_pro = f"{numero_poliza_birlik}_{numero_documento}"
                                factura = obtener_ultimo_archivo_descargado_x_identificador(id_prof_pro,ruta_carpeta_facturas)
                                codCuota, codGiro = obtener_cod_cuota(factura)
                                print(f"✅ Código de cuota extraído: {codCuota}, Código de Giro extraído: {codGiro}")
                            
                                try:

                                    # Caso 1: Coincidencia con codCuota
                                    if numero_proforma_birlik == codCuota or numero_proforma_birlik == codGiro:

                                        print("✅ Uno de los codigos coinciden con el de Birlik.")
                                        fila_encontrada_codCuota = True
                                        codigo_giro.append(codCuota)
                                        codigo_giro.append(codGiro)
                                        #fechas_factura = datetime.strptime(fecha_emision, "%d/%m/%Y")
                                        fechas_factura = fecha_emision
                                        break

                                    else:
                                        continue

                                finally:
                                    os.remove(ruta_factura)

                        # Si la fila se encontró → salir del while
                        if fila_encontrada_codCuota:
                            break #*(While)

                        # --- Si NO se encontró fila → intentar pasar de página ---
                        print("➡ Pasando a la siguiente página de la Tabla")

                        #time.sleep(5)
                        
                        try:
                            # Buscar botones habilitados
                            botones = driver.find_elements(By.CSS_SELECTOR,"button.pagination-step-previous:not([disabled])")

                            # Filtrar el que tenga el path que empiece con M1.5
                            btn_siguiente = None
                            for b in botones:
                                try:
                                    path = b.find_element(By.TAG_NAME, "path")
                                    if path.get_attribute("d").startswith("M1.5"):
                                        btn_siguiente = b
                                        break
                                except:
                                    continue

                            if not btn_siguiente:
                                print("❌ No hay más páginas")
                                break

                            # Esperar click real
                            wait.until(EC.element_to_be_clickable(btn_siguiente))
                            # Clic
                            btn_siguiente.click()
                            print("➡ Clic en botón SIGUIENTE")

                        except:
                            print("❌ No hay más páginas")
                            break

                        # esperar a que la tabla recargue
                        time.sleep(4)
                    
                    if not fila_encontrada_codCuota:
                        raise Exception (f"No se encontró el documento en ninguna página.")

                    break #(***)

            if not fila_encontrada:
                raise Exception (f"No se encontró ninguna fila con la Poliza: '{numero_poliza_birlik}'.")
   
    except Exception as e:
            print(f"❌ Error Procesando toda la fila para Salud, Motivo: {e}")
    finally:
            return codigo_giro,fechas_factura

def obtener_ultimo_archivo_descargado_x_identificador(identificador,ruta_carpeta_facturas):

    patron = os.path.join(ruta_carpeta_facturas, f"{identificador}.pdf")
    archivos = glob.glob(patron)

    if not archivos:
        raise FileNotFoundError(f"No se encontró el documento {identificador}.pdf")

    # Ordenar por fecha de modificación (último primero)
    archivos.sort(key=os.path.getmtime, reverse=True)
    return archivos[0]  # El más reciente

def obtener_cod_cuota(factura_pdf):
    
    texto_total = ""

    # Abrir PDF y extraer todo el texto
    with pdfplumber.open(factura_pdf) as pdf:
        for pagina in pdf.pages:
            texto_total += pagina.extract_text() + "\n"

    # # Buscar solo el campo Documento:
    # match_documento = re.search(r"Documento\s*:\s*(.+)", texto_total, re.IGNORECASE)
    # # Si lo encuentra, devolver tal cual (sin convertir a minúscula)
    # cod_cuota = match_documento.group(1).strip() if match_documento else None

    match_documento = re.search(r"Documento\s*:\s*.*-(\d+)", texto_total, re.IGNORECASE)
    cod_cuota = match_documento.group(1) if match_documento else None

    # Buscar C-xxxxx (solo una vez)
    match_c = re.search(r"C-(\w+)", texto_total)
    cod_giro = match_c.group(1) if match_c else None

    return cod_cuota, cod_giro

def main():
    
    display_num = os.getenv("DISPLAY_NUM", "0")
    os.environ["DISPLAY"] = f":{display_num}"

    driver, wait = abrirDriver()

    # ------------------ Inicio del Flujo de Automatización ------------------
    driver.get(urlPacifico)
    print("✅ Ingresando a la URL.")
           
    mi_portafolio = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Mi portafolio')]")))
    mi_portafolio.click()
    print("🖱️ Clic en Portafolio.")

    somos_corredores = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Somos Corredores')]")))
    somos_corredores.click()
    print("🖱️ Clic en Somos Corredores")

    user_input = wait.until(EC.element_to_be_clickable((By.ID, "i0116")))
    user_input.clear()
    user_input.send_keys(correo)
    print("⌨️ Digitando el Correo")

    boton_next = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
    boton_next.click()
    print("🖱️ Clic en 'Next'.")
        
    pass_input = wait.until(EC.element_to_be_clickable((By.ID, "i0118")))
    pass_input.clear()
    pass_input.send_keys(password)
    print("⌨️ Digitando el Password")

    ingresar_btn = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
    ingresar_btn.click()
    print("🖱️ Clic en 'Ingresar'.")

    sms_option = wait.until(EC.element_to_be_clickable((By.XPATH,"//div[@class='table' and @data-value='OneWaySMS']")))
    driver.execute_script("arguments[0].click();", sms_option)
    print("🖱️ Clic en 'Enviar un mensaje de texto'.")

    codigo_path = "/codigo/codigo.txt"

    print("⏳ Esperando código...")
    while not os.path.exists(codigo_path):
        time.sleep(2)

    with open(codigo_path, "r") as f:
        codigo = f.read().strip()

    print(f"✅ Código recibido desde volumen: {codigo}")

    clave_sms = wait.until(EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCC_OTC")))
    clave_sms.clear()
    clave_sms.send_keys(codigo)
    print("⌨️ Digitando el código")

    # --- Eliminar el archivo después de usarlo ---
    try:
        os.remove(codigo_path)
        #print("🧹 Archivo codigo.txt eliminado del volumen.")
    except FileNotFoundError:
        print("⚠️ No se encontró codigo.txt al intentar eliminarlo (ya fue borrado).")
    except Exception as e:
        print(f"❌ Error al eliminar codigo.txt: {e}")

    ingresar_btn = wait.until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCC_Continue")))
    ingresar_btn.click()
    print("🖱️ Clic en 'Ingresar'.")

    boton_conf = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
    boton_conf.click()
    print("🖱️ Clic en 'Yes'.")

    while True:

        try:
            todos_los_datos = consultarAPI(url_datos_para_cancelar_cuotas,ids_compania)

            if not todos_los_datos:
                raise Exception("No se recibió información de ninguna compañía.")

            log_path,ruta_salida_API,ruta_salida,ruta_maestro,nombre_hoja, ruta_carpeta_facturas, ruta_carpeta_comprobante, ruta_carpeta_errores,carpeta_compañia,carpeta_principal = crearCarpetas(todos_los_datos,nombre_carpeta_compañia,tipo=2,cia_a_verificar=None)

            print("\n📁 Iniciando procesamiento para Pacifico...")

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
