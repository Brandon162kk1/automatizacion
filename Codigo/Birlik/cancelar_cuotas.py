#-- Imports ---
import time
import os
#-- Froms ----
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException
from datetime import datetime
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait

# Fecha de Inicio de Birlik 
fecha_creacion_birlik = datetime(2022,12,5)
#-------BIRLIK ENLACES-------
url_datos_para_enviar_factura = os.getenv("url_datos_para_enviar_factura")

url_datos_para_cancelar_cuotas = os.getenv("url_datos_para_cancelar_cuotas")
url_historial_cuotas = os.getenv("url_historial_cuotas")
#Historial de cuotas (pendientes y canceladas) pero no se puede anular aqui
url_cuotas_canceladas = os.getenv("url_cuotas_canceladas")
# Hisorial de cuotas canceladas, listar para enviar facturas por mayor
url_cuotas = os.getenv("url_cuotas")
# Mini historial de cuotas que si se pueden anular
url_detalle_poliza = os.getenv("url_detalle_poliza") 
# Pagina Detalle de la Poliza, se puede enviar factura y mensaje de cobranza, pero no es recomendable (falta ajustar)
url_agregar_comprobante = os.getenv("url_agregar_comprobante")
url_para_cobrar_cuotas = os.getenv("url_para_cobrar_cuotas")
#--------BIRLIK-------------
login_birlik = os.getenv("login_birlik") #Cancelar Cuotas
usuarioBirlik = os.getenv("usuarioBirlik")
passwordBirlik = os.getenv("passwordBirlik")

# Se sube la factura como comprobante de pago, 'CodigoCuota_Comprobante_dia_mes'
def agregar_comprobante_pago(driver, wait, id_cuota_birlik,ruta_factura):
    
    # Guardar ventana principal (CIA)
    ventana_principal_cia = driver.current_window_handle

    try:

        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[-1])

        # Construir la URL completa con el idCuota
        url_final = url_agregar_comprobante + id_cuota_birlik
        driver.get(url_final)
        
        try:
            # Intentar encontrar los campos de login
            email_input = WebDriverWait(driver,5).until(EC.presence_of_element_located((By.ID, 'signinSrEmail')))
            password_input = WebDriverWait(driver,5).until(EC.presence_of_element_located((By.ID, 'signupSrPassword')))

            # Si llega aquí, significa que está en la pantalla de login
            email_input.clear()
            email_input.send_keys(usuarioBirlik)

            password_input.clear()
            password_input.send_keys(passwordBirlik)
            password_input.send_keys(Keys.RETURN)

            print("🔑 Sesión iniciada en Birlik.")

        except TimeoutException:
            # Si no encontró los elementos, asumimos que ya está logueado
            print("✅ Sesión ya activa, no es necesario iniciar.")

        time.sleep(5)

        archivo_comprobante = wait.until(EC.presence_of_element_located((By.ID, "customFile_comprobante")))
        archivo_comprobante.send_keys(ruta_factura)
        print("✔ Comprobante de Pago cargado.")

        boton_adjuntar = driver.find_element(By.XPATH, '//button[contains(text(), "ADJUNTAR")]')
        boton_adjuntar.click()
        print("🟢 Se hizo clic en 'ADJUNTAR' para cancelar la cuota.")

        time.sleep(5)
        print("Proceso de cancelación y acceso a Birlik completado.")

        return True

    except StaleElementReferenceException as e:
        print("⚠ DOM modificado tras registrar. No es un error grave. Detalles:", e)
        return False
    except Exception as e:
        print(f"❌ Error durante el registro del comprobante en Birlik. Error: {e}")
        return False
    finally:
        driver.close()
        driver.switch_to.window(ventana_principal_cia)
        print("🔄 Cerrando Birlik y volviendo a la CIA")

# Se sube la Constancia de la SUNAT  y la Factura
def cancelar_y_agregar_cuota(driver, wait, id_cuota,comprobante_valor,fecha_emision,ruta_factura,ruta_imagen_sunat,resultado_importe):
    
    # Guardar ventana principal (CIA)
    ventana_principal_cia = driver.current_window_handle

    try:

        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[-1])

        # Construir la URL completa con el idCuota
        url_final = login_birlik + id_cuota
        driver.get(url_final)
       
        try:
            # Intentar encontrar los campos de login
            email_input = WebDriverWait(driver,5).until(EC.presence_of_element_located((By.ID, 'signinSrEmail')))
            password_input = WebDriverWait(driver,5).until(EC.presence_of_element_located((By.ID, 'signupSrPassword')))

            # Si llega aquí, significa que está en la pantalla de login
            email_input.clear()
            email_input.send_keys(usuarioBirlik)

            password_input.clear()
            password_input.send_keys(passwordBirlik)
            password_input.send_keys(Keys.RETURN)

            print("🔑 Sesión iniciada en Birlik.")

        except TimeoutException:
            # Si no encontró los elementos, asumimos que ya está logueado
            print("✅ Sesión ya activa, no es necesario iniciar.")
        
        time.sleep(5)

        # Rellenar campos del formulario   
        factura_input = wait.until(EC.presence_of_element_located((By.ID, "factura")))
        factura_input.clear()
        factura_input.send_keys(ruta_factura)
        print("✅ Factura subida.")
        
        constancia_input = wait.until(EC.presence_of_element_located((By.ID, "constancia")))
        constancia_input.clear()
        constancia_input.send_keys(ruta_imagen_sunat)
        #print(f"Ruta de la constancia: {ruta_imagen_sunat} ")
        print("✅ Constancia de SUNAT subida.")

        archivo_input_1 = wait.until(EC.presence_of_element_located((By.ID, "comprobante")))
        archivo_input_1.send_keys(ruta_factura)
        print("✅ Comprobante subido.")

        observacion_input = wait.until(EC.presence_of_element_located((By.ID, "observacionCuota")))
        observacion_input.clear()
        if not resultado_importe:
            observacion_input.send_keys("Importes no coinciden")
        else:
            observacion_input.send_keys("Cancelado por robot")
        print("✅ Observación ingresada.")

        # Número de factura
        input_factura = wait.until(EC.presence_of_element_located((By.ID, "numeroFactura")))
        input_factura.clear()
        input_factura.send_keys(comprobante_valor)
        print("🧾 Número de factura ingresado:", comprobante_valor)

        # Ojo BIRLIK espera el formato MM/DD/YYYY
        # Paso 1: Convertir de "24/05/2025" a datetime -- fecha_emision_probar = datetime.strptime(fecha_emision_valor, "%d/%m/%Y")
        fecha_dt = datetime.strptime(fecha_emision, "%d/%m/%Y")
        fecha_vista_factura = fecha_dt.strftime("%d/%m/%Y") #Asi se ve en la factura , sin hora ni minutos
        fecha_formateada_js  = fecha_dt.strftime("%Y-%m-%d")  # 👈 Formato compatible con <input type="date">

        input_fecha = wait.until(EC.presence_of_element_located((By.ID, "fechaPago")))
        input_fecha.clear()
        # Usar JavaScript para asignar el valor
        driver.execute_script("arguments[0].value = arguments[1];", input_fecha, fecha_formateada_js)
        print(f"📅 Fecha de Emisión de la Factura: {fecha_vista_factura}, pero se ingresa así (Y/m/d): {fecha_formateada_js}")

        btn_registrar = wait.until(EC.element_to_be_clickable((By.ID, "btnRegistrar")))
        btn_registrar.click()
        print("🖱️ Clic en 'REGISTRAR' para cancelar la cuota.")

        # 3. Esperar a que cargue la página de confirmación
        wait.until(EC.invisibility_of_element_located((By.ID, "btnRegistrar")))
        print("✅ Confirmación de registro detectada.")

        time.sleep(5)

        return True

    except StaleElementReferenceException as e:
        print("⚠ DOM modificado tras registrar. No es un error grave. Detalles:", e)
        return False
    except Exception as e:
        print(f"❌ Error durante el registro en Birlik. Detalles: {e}")
        return False
    finally:
        # Cerrar la pestaña de Birlik
        driver.close()
        driver.switch_to.window(ventana_principal_cia)
        print("🔄 Cerrando Birlik y volviendo a la CIA")

# Se cobra la cuota a los clientes por medio de Birlik
def cobrarCuota(driver,wait,id_cliente,codigoCuota):

    try:
        print("Entrando a Birlik para mandar correo de cobranza")
        url_final = url_cuotas + id_cliente
        driver.get(url_final)

        # # Ingresar credenciales
        # email_input = wait.until(EC.presence_of_element_located((By.ID, 'signinSrEmail')))
        # email_input.clear()
        # email_input.send_keys(username_birlik)

        # password_input = wait.until(EC.presence_of_element_located((By.ID, 'signupSrPassword')))
        # password_input.clear()
        # password_input.send_keys(password_birlik)
        # password_input.send_keys(Keys.RETURN)

        # print("Iniciando sesión en Birlik...")
        # time.sleep(5)

        # Buscar el input por ID y escribir texto
        input_busqueda = wait.until(EC.presence_of_element_located((By.ID, "buscar_cuotas")))
        #driver.find_element(By.ID, "buscar_cuotas")
        input_busqueda.clear()
        input_busqueda.send_keys(codigoCuota)
        print(f"🔎 Buscando la Cuota: {codigoCuota}")

        # Espera e identifica la tabla y buscar la fila donde la columna "Codigo" contenga el codigo de Cuota.
        table = wait.until(EC.presence_of_element_located((By.ID, "tb_detallecuotas")))

        rows = table.find_elements(By.TAG_NAME, "tr")

        # # Buscar el input por ID y escribir texto
        # input_busqueda = wait.until(EC.presence_of_element_located((By.ID, "buscar_cuotas")))
        # #driver.find_element(By.ID, "buscar_cuotas")
        # input_busqueda.clear()
        # input_busqueda.send_keys(codigoCuota)
        # print(f"Buscando la Cuota: {codigoCuota}")

        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) >= 5:  # Asegurarse que es una fila de 6 celdas
                print("✅ Codigo de cuota encontrando")
                elemento = driver.find_element(By.CSS_SELECTOR, 'a[onclick="enviarmensajec(9933, this)"]')
                elemento.click()
                print("🖱️ Clic en el icono mensaje")
                print("⌛ Cargando el modal de cobranza")
                # 1. Esperar a que aparezca el input con el correo (asegura que el modal se cargó)
                wait.until(EC.visibility_of_element_located((By.ID, "correoClienteInput2")))
                print("⌛ Esperar el boton Mensaje")
                # 2. Esperar el botón "Enviar Mensaje"

                # Esperar a que el botón esté presente en el DOM
                boton_enviar = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//button[@type='submit' and contains(@onclick, 'enviarYAnimar2')]"))
                )

                # Hacer scroll hasta el botón
                driver.execute_script("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", boton_enviar)

                boton_enviar.click()

                # boton_enviar = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@onclick="enviarYAnimar2()"]')))
                # print("Listo para enviar")
                # # 3. Hacer clic en el botón
                # boton_enviar.click()
                print("🖱️ Clic en Enviar")
                time.sleep (10)

        #driver.quit()
        return # ✅ <--- Este return es obligatorio para cerrar el flujo correctamente

    except Exception as e:
        print(f"❌ Error durante el envio de la Factura al cliente por medio de Birlik, {e}")
    # finally:
    #     driver.quit()

def main():
    mensaje = "Hola"
    print(mensaje)

if __name__ == "__main__":
    main()