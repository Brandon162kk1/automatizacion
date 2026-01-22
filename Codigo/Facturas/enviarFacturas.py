#-- Imports --
import os
import time
import pandas as pd
#-- Froms --
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from collections import defaultdict
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from Birlik.cancelar_cuotas import url_cuotas_canceladas,usuarioBirlik,passwordBirlik,url_datos_para_enviar_factura
from Apis.api_birlik import consultarAPI
from GoogleChrome.chromeDriver import crearCarpetas,abrirDriver
from GoogleChrome.fecha_y_hora import get_dia,get_mes,get_anio

def analizarFacturasparaEnviarCliente(ruta_maestro_evaluar):
    try:
        # Leer directamente la única hoja (lo que guardaste del API)
        df = pd.read_excel(ruta_maestro_evaluar, engine="openpyxl")
    except Exception as e:
        print(f"❌ Error al abrir el Excel de las cuotas para enviar Facturas, Detalle: {e}")
        return

    # ---------------- AGRUPACIÓN POR CLIENTE ----------------
    clientes_codigos = {}
    for _, row in df.iterrows():
        tipodoc = str(row['tipoDocumento']).strip()

        datos_cuota = {
            'id_poliza': row['id_Poliza'],
            'fk_ramo': row['fk_Ramo'],
            'codigoCuota': row['codigoCuota'],
            'fk_Cliente': row['fk_Cliente'],
            'id_Cuota': row['id_Cuota'],
            'asegurado': row['asegurado']
        }

        clientes_codigos.setdefault(tipodoc, []).append(datos_cuota)

    # ---------------- ENVIAR POR CLIENTE ----------------
    for tipo, lista_cuotas in clientes_codigos.items():

        display_num = os.getenv("DISPLAY_NUM", "0")  # fallback = 0
        os.environ["DISPLAY"] = f":{display_num}"

        print(f"--- Enviando Facturas el {get_dia()}-{get_mes()}-{get_anio()} ---")

        driver, wait = abrirDriver() #ruta_carpeta_descargas

        try:
            enviarFacturasCliente(driver, wait, lista_cuotas)
        except Exception as e:
            print(f"❌ Error enviando facturas para {tipo}: {e}")
        finally:
            driver.quit()

def enviarFacturasCliente(driver, wait, lista_cuotas):

    # Agrupar cuotas por cliente
    grupos, nombre_por_cliente = agrupar_por_cliente(lista_cuotas)

    # Ir a una URL que requiera login y loguear una sola vez
    primer_fk = next(iter(grupos.keys()))
    driver.get(f"{url_cuotas_canceladas}{primer_fk}")
    login_un_avez(wait, usuarioBirlik, passwordBirlik)

    # Recorrer cada cliente y enviar sus cuotas
    for fk_cliente, cuotas in grupos.items():
        nombre_asegurado = nombre_por_cliente.get(fk_cliente, "(sin nombre)")
        url_final = f"{url_cuotas_canceladas}{fk_cliente}"
        driver.get(url_final)
        print(f"🌐 Cliente {nombre_asegurado}")

        # Marcar las cuotas de ESTE cliente
        for q in cuotas:
            # Asegura strings si tus helpers comparan texto
            codigo = str(q.get('codigoCuota', '')).strip()
            id_cuota = str(q.get('id_Cuota', '')).strip()
            if not codigo or not id_cuota:
                print(f"⚠️ Cuota inválida (sin código o id): {q}")
                continue

            buscar_y_seleccionar_checkbox(driver, wait, codigo, id_cuota)

        # Enviar una sola vez para todas las seleccionadas de ese cliente
        if clic_enviar_mensaje(driver, wait):
            print(f"❌ Plataforma Birilk en Mantenimiento")
        else:
            print(f"✅ Envío realizado para {nombre_asegurado} (ID {fk_cliente})")
        print("--------------------------------------------------")

def agrupar_por_cliente(lista_cuotas):
    grupos = defaultdict(list)
    nombre_por_cliente = {}
    for q in lista_cuotas:
        fk = q.get('fk_Cliente')
        if fk is None:
            print(f"⚠️ Cuota sin fk_Cliente, se omite: {q}")
            continue
        grupos[fk].append(q)
        nom = (q.get('asegurado') or "").strip()
        # Guarda el primero no-vacío que encuentres
        if fk not in nombre_por_cliente and nom:
            nombre_por_cliente[fk] = nom
    return grupos, nombre_por_cliente

def login_un_avez(wait, username, password):

    try:
        email_input = wait.until(EC.presence_of_element_located((By.ID, 'signinSrEmail')))
        email_input.clear()
        email_input.send_keys(username)
        print("✅ Ingresando Usuario")

        time.sleep(2)

        password_input = wait.until(EC.presence_of_element_located((By.ID, 'signupSrPassword')))
        password_input.clear()
        password_input.send_keys(password)
        print("✅ Ingresando Contraseña")

        password_input.send_keys(Keys.RETURN)
        time.sleep(5)

        print("🔐 Sesión iniciada")
    except TimeoutException:
        # Si no aparecen campos, probablemente ya estás logueado
        print("ℹ️ Ya estabas logueado o no se requirió login.")

def buscar_y_seleccionar_checkbox(driver, wait, codigoCuota, id_Cuota):

    try:
        input_busqueda = wait.until(EC.presence_of_element_located((By.ID, "buscar_cuotas")))
        input_busqueda.clear()
        input_busqueda.send_keys(codigoCuota)   
        time.sleep(2)

        label_for_id = f"cb-{id_Cuota}"
        label_selector = f"label[for='{label_for_id}']"

        label = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, label_selector)))

        driver.execute_script("arguments[0].scrollIntoView(true);", label)
        time.sleep(0.3)
        label.click()

        print(f"✅ Cuota {codigoCuota} seleccionada correctamente")
    except Exception as e:
        print(f"❌ Error seleccionando checkbox de {codigoCuota}: {e}")

def clic_enviar_mensaje(driver, wait):
    boton = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@onclick="enviarmensaje(this)"]')))
    boton.click()
    print("🖱️ Clic en 'Enviar Mensaje' ")

    time.sleep(3)

    #wait.until(EC.visibility_of_element_located((By.ID, 'correoClienteInput2')))
    boton_enviar_modal = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@type="submit" and contains(text(), "Enviar Mensaje")]')))
    driver.execute_script("arguments[0].scrollIntoView();", boton_enviar_modal)

    time.sleep(3)

    boton_enviar_modal.click()
    print("🖱️ Clic en 'Enviar' ")

    try:
        WebDriverWait(driver,5).until(EC.presence_of_element_located((By.XPATH,"//h1[normalize-space()='¡Hola! Lamentamos la interrupción.']")))
        return True
    except TimeoutException:
        return False

def main():
    
    tipo = 0 #--> momentaneo para enviar facturas a los clientes
    #ids_compania = [23,26,5,29,31,15,32,2,35,38,1,18,27,24,4,36,17,33,39,16,25,11,12,13,14]
    # A un futuro traes todas las compañias que nos generen cuotas para enviar facturas
    ids_compania = [i for i in range(1, 39) if i != 37]
    json_vacio = {}

    while True:

        data_facturas_por_enviar = consultarAPI(url_datos_para_enviar_factura,ids_compania)

        if data_facturas_por_enviar:
            nom_car_pri= f"Facturas_Enviadas_{get_dia()}_{get_mes()}"
            ruta_salida_facturas, log_path = crearCarpetas(json_vacio, nom_car_pri, tipo, cia_a_verificar=None)

            df_final = pd.concat(data_facturas_por_enviar, ignore_index=True)
            df_final.to_excel(ruta_salida_facturas, index=False)

            time.sleep(1)
            analizarFacturasparaEnviarCliente(ruta_salida_facturas)
            time.sleep(1)
     
            if os.path.exists(ruta_salida_facturas):
                os.remove(ruta_salida_facturas)
                #print("Archivo .xlsx eliminado correctamente")
            else:
                print("⚠️ La carpeta no existe")

        #print("⌛ Esperando 3 segundos antes de revisar nuevamente...\n")
        time.sleep(3)

if __name__ == "__main__":
    main()
