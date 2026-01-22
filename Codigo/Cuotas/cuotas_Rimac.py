#-- Imports ---
import re
import sys
import subprocess
import time
import os
import pandas as pd
import shutil
import requests
import urllib.parse as urlparse
#import pyautogui  # Solo si necesitas automatizar ventanas nativas
#-- Froms ----
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from Excels.estilosExcel import guardar_excel_con_formato
from Sunat.validar_factura import consultarValidezSunat,login_sunat
from Birlik.cancelar_cuotas import agregar_comprobante_pago,cancelar_y_agregar_cuota,url_cuotas,url_cuotas_canceladas,url_datos_para_cancelar_cuotas
from Apis.api_birlik import consultarAPI
from GoogleChrome.chromeDriver import abrirDriver, crearCarpetas
from GoogleChrome.fecha_y_hora import get_timestamp
#from correoit.correo_ariadne import revisar_correo_ariadne

#--------- COMPAÑÍA RIMAC ------
# Lista de IDs de compañía
ids_compania = [27,28,35]
#----- Variables de Entorno -------
urlRimacCorredores = os.getenv("urlRimacCorredores")
usernameRimacCorredores = os.getenv("remitente")
passwordRimacCorredores = os.getenv("passwordCorredores")
#----- Carpeta de la Compañia -------
nombre_carpeta_compañia = f"Rimac_{get_timestamp()}"

# --- Configuración ---
API_KEY = "8b61feec172173ef48060a723af1b6c7"

def resolver_recaptcha(driver, wait, API_KEY):

    try:
        iframe = wait.until(EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, 'reCAPTCHA')]")))
        src = iframe.get_attribute("src")
        parsed = urlparse.urlparse(src)
        sitekey = urlparse.parse_qs(parsed.query)["k"][0]

        print("🔑 Sitekey:", sitekey)

        resp = requests.post("http://2captcha.com/in.php", data={
            "key": API_KEY,
            "method": "userrecaptcha",
            "googlekey": sitekey,
            "pageurl": driver.current_url
        })
        if not resp.text.startswith("OK|"):
            raise Exception("2Captcha error: " + resp.text)

        captcha_id = resp.text.split("|")[1]
        print("⌛ Esperando solución:", captcha_id)

        token = None
        for i in range(30):
            time.sleep(5)
            r = requests.get(f"http://2captcha.com/res.php?key={API_KEY}&action=get&id={captcha_id}")
            if "OK|" in r.text:
                token = r.text.split("|")[1]
                break

        if not token:
            raise Exception("Timeout esperando 2Captcha")

        driver.execute_script(
            'document.getElementById("g-recaptcha-response").value = arguments[0];',
            token
        )
        driver.execute_script("""
            document.getElementById("g-recaptcha-response").dispatchEvent(new Event('change'));
        """)

        #print("✅ reCAPTCHA resuelto y aplicado")
        return True

    except Exception as e :
        print(e)
        return False

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

    resultado_importe = False
    resultado_sunat = False
    resultado_birlik = False
    resultado_estado = None
    resultado_accion = ""

    try:

        # ------------------ Inicio del Flujo de Automatización ------------------

        dropdown_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.dropdown_dropdown__icon__lTwKq")))
        dropdown_btn.click()
        print("Clic en Dropdown")

        opcion_tipo_documento = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Tipo de documento']")))
        opcion_tipo_documento.click()
        print("Clic en tipo de documento")

        dropdown2_btn = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "div.dropdown_dropdown__main__KksvY button.dropdown_dropdown__icon__lTwKq")
        ))
        dropdown2_btn.click()
        print("Clic en tipo de documento")

        if tipo_doc_birlik == 'DNI':
            texto_opcion = 'D.N.I.'
        elif tipo_doc_birlik == 'CEX':
            texto_opcion = 'CARNET DE EXTRANJERIA'
        else:
            texto_opcion = 'R.U.C.'

        opcion = wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[normalize-space()='{texto_opcion}']")))
        opcion.click()
        print(f"Clic en {texto_opcion}")

        campo_documento = wait.until(EC.presence_of_element_located((By.ID, "documentNumber")))
        campo_documento.clear()
        campo_documento.send_keys(ruc_cliente_birlik)
        print(f"✍️ Número de documento ingresado: {ruc_cliente_birlik}")

        input("Esperar")

    except Exception as e:
        print(f"❌ Error Procesando toda la fila, Revisar: {e}")
    finally:
        return f"Coinciden" if resultado_importe else f"No coinciden" ,"Válido" if resultado_sunat else "No Válido" ,"Cuota Cancelada" if resultado_birlik else "Cuota Pendiente" ,resultado_estado,resultado_accion

def main():
  
    display_num = os.getenv("DISPLAY_NUM", "0")
    os.environ["DISPLAY"] = f":{display_num}"

    driver, wait = abrirDriver()#carpeta_compañia

    driver.get(urlRimacCorredores) 
    print("⌛ Ingresando a la URL.")

    user_input = wait.until(EC.presence_of_element_located((By.ID, ":ride:1")))
    user_input.clear()
    user_input.send_keys(usernameRimacCorredores)
    print("⌨️ Digitando el Correo.")
        
    pass_input = wait.until(EC.presence_of_element_located((By.ID, ":ride:2")))
    pass_input.clear()
    pass_input.send_keys(passwordRimacCorredores)
    print("⌨️ Digitando el Password")

    ingresar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and .//span[text()='Ingresar']]")))
    driver.execute_script("arguments[0].click();", ingresar_btn)
    print("🖱️ Clic en 'Ingresar'.")

    # try:
    #     if resolver_recaptcha(driver, wait, API_KEY):
    #         print("✅ reCAPTCHA resuelto y aplicado")
    #     else:
    #         input("Esperar")
    # except TimeoutException:
    #     print ("No salio")

    codigo_path = "/codigo_rimac/codigo.txt"

    print("⏳ Esperando código...")
    while not os.path.exists(codigo_path):
        time.sleep(2)

    with open(codigo_path, "r") as f:
        codigo = f.read().strip()

    print(f"✅ Código recibido desde volumen: {codigo}")

    inputs = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input._input_12twg_6[maxlength='1']")))

    if len(inputs) == len(codigo):
        for i, inp in enumerate(inputs):
            inp.clear()
            inp.send_keys(codigo[i])
    else:
        print(f"⚠️ Inputs encontrados: {len(inputs)}, dígitos del código: {len(codigo)}")

    ingresar_btn2 = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Ingresar']]")))
    driver.execute_script("arguments[0].click();", ingresar_btn2)
    print("🖱️ Clic en 'Ingresar'.")

    cobranza = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='Cobranza']")))
    try:
        cobranza.click()
        print("Clic normal")
    except:
        driver.execute_script("arguments[0].click();", cobranza)
        print("Clic fuerza")

    revisar_pagos = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='Revisar pagos']")))
    try:
        revisar_pagos.click()
        print("Clic normal")
    except:
        driver.execute_script("arguments[0].click();", revisar_pagos)
        print("Clic fuerza")

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

            guardar_excel_con_formato(ruta_salida,'Sheet1')

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
