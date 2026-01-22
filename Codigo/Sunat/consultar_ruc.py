import time
import os
import sys
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException
from GoogleChrome.chromeDriver import crearCarpetas, abrirDriver
from Excels.estilosExcel import guardar_excel_con_formato_solo_ajustar_columnas
from GoogleChrome.fecha_y_hora import get_dia,get_mes

nom_car_pri = f'Actividades_Clientes_{get_dia()}_{get_mes()}'
url_consultar_ruc = 'https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp'

def main():
    
    json_vacio = {}
    ruta_salida_facturas,log_path ,carpeta_principal = crearCarpetas(json_vacio ,nom_car_pri,tipo=5,cia_a_verificar = None)
    # --- Construir ruta para el log ---
    base_dir = os.path.dirname(os.path.abspath(__file__))

    # # --- Redireccionar todo el stdout al archivo de log ---
    # original_stdout = sys.stdout  # Guarda referencia original a la consola
    # with open(log_path, "w", encoding="utf-8") as log_file:

    #     sys.stdout = Tee(sys.stdout, log_file)

    display_num = os.getenv("DISPLAY_NUM", "0")  # fallback = 0
    os.environ["DISPLAY"] = f":{display_num}"

    driver, wait = abrirDriver(carpeta_principal)

    driver.get(url_consultar_ruc)
    print("✅ Ingresando a la URL de SUNAT")

    try:
        # Aquí va TODO tu código que imprime cosas
        print("📁 Iniciando procesamiento...")           

        nombre_salida_API = "CLIENTESRUC.xlsx"
        # 1. Leer el Excel con pandas
        ruta_excel = os.path.join(base_dir, nombre_salida_API)

        try:
            df = pd.read_excel(ruta_excel, engine="openpyxl", dtype={"NumeroDocumento": str})
        except Exception as e:
            print(f"❌ Error al leer el archivo Excel: {e}")
            return

        # Inicializa columnas si no existen
        df["Cod_Princiapl"] = ""
        df["Actividad_Principal"] = ""
        df["Cod_Secundario_1"] = ""
        df["Actividad_1"] = ""
        df["Cod_Secundario_2"] = ""
        df["Actividad_2"] = ""

        total_filas = len(df)

        # 2. Iterar sobre cada fila
        for index, row in df.iterrows():
            print(f"\n--- Procesando fila {index + 2} de {total_filas + 1} ---")
            try:
                cod_pri,act_prin,cod_sec_1 ,act_sec_1,cod_sec_2,act_sec_2 = procesar_fila(row,driver,wait,carpeta_principal)

                df.at[index, "Cod_Princiapl"] = cod_pri
                df.at[index, "Actividad_Principal"] = act_prin
                df.at[index, "Cod_Secundario_1"] = cod_sec_1
                df.at[index, "Actividad_1"] = act_sec_1
                df.at[index, "Cod_Secundario_2"] = cod_sec_2
                df.at[index, "Actividad_2"] = act_sec_2

                # 💾 Guardar luego de procesar cada fila
                df.to_excel(ruta_salida_facturas, index=False)

                print(f"✔ Fila {index} guardada correctamente.")

            except Exception as e:
                print(f"❌ Error en fila {index}, {e}")
                df.to_excel(ruta_salida_facturas, index=False)
                continue

            time.sleep(2)

        # Guardar el DataFrame en un nuevo archivo Excel
        try:
            df.to_excel(ruta_salida_facturas, index=False)
            print(f"\n✅ Validaciones completadas. Archivo guardado en:\n{ruta_salida_facturas}")
        except Exception as e:
            print(f"\n❌ Error al guardar el archivo Excel: {e}")

        print(f"\n✅ Proceso finalizado. Todo el log ha sido guardado en:\n{log_path}")

        guardar_excel_con_formato_solo_ajustar_columnas(ruta_salida_facturas,'Sheet1')

    finally:
        # sys.stdout = original_stdout  # Restaura la salida estándar (la consola)
        pass
   
def procesar_fila(row,driver,wait,carpeta_principal):
    
    cliente = str(row["CLIENTE"]).strip()
    ruc = str(row["RUC"]).strip()

    cod_pri = ""
    act_prin = ""
    cod_sec_1 = ""
    act_sec_1 = ""
    cod_sec_2 = ""
    act_sec_2 = ""

    try:

        campo_ruc_rec = wait.until(EC.element_to_be_clickable((By.NAME, "search1")))
        campo_ruc_rec.clear()
        campo_ruc_rec.send_keys(ruc)
        print(f" Ruc ingresado: {ruc} para el cliente {cliente}")

        time.sleep(2)

        buscar_btn = wait.until(EC.element_to_be_clickable((By.ID, "btnAceptar")))
        buscar_btn.click()
        print("🖱️ Clic en Buscar")

        time.sleep(2)

        try:

            # Esperar hasta que aparezca el bloque de Actividad Económica
            bloque_actividad = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//h4[contains(text(),'Actividad(es) Económica(s):')]/ancestor::div[@class='row']")
                )
            )

            #nombre_archivo = f"{cliente}_{ruc}.png"
            #ruta_archivo = os.path.join(carpeta_principal, nombre_archivo)
            #driver.save_screenshot(ruta_archivo)

            #print(f"📸 Captura guardada")

            # Dentro del bloque buscar la tabla
            tabla = bloque_actividad.find_element(By.XPATH, ".//table[@class='table tblResultado']")

            # Extraer todas las filas
            filas = tabla.find_elements(By.TAG_NAME, "tr")

            actividades = {}  # aquí vamos a guardar los resultados

            for fila in filas:
                texto = fila.text.strip()
                partes = [p.strip() for p in texto.split("-")]
                if len(partes) >= 3:
                    tipo = partes[0]   # Principal / Secundaria 1 / Secundaria 2
                    codigo = partes[1]
                    actividad = partes[2]
                    actividades[tipo] = {
                        "codigo": codigo,
                        "actividad": actividad
                    }

            # Mapear a variables individuales
            if "Principal" in actividades:
                cod_pri = actividades["Principal"]["codigo"]
                act_prin = actividades["Principal"]["actividad"]

            if "Secundaria 1" in actividades:
                cod_sec_1 = actividades["Secundaria 1"]["codigo"]
                act_sec_1 = actividades["Secundaria 1"]["actividad"]

            if "Secundaria 2" in actividades:
                cod_sec_2 = actividades["Secundaria 2"]["codigo"]
                act_sec_2 = actividades["Secundaria 2"]["actividad"]

            # Debug en consola
            print(f"✅ Principal: {cod_pri} | {act_prin}")
            print(f"✅ Secundaria 1: {cod_sec_1} | {act_sec_1}")
            print(f"✅ Secundaria 2: {cod_sec_2} | {act_sec_2}")

        except Exception as e:
            print(f"❌ Error al obtener actividades económicas: {e}")

        time.sleep(2)

        try:
            # Intentar con el botón visible
            boton_volver = driver.find_element(By.CSS_SELECTOR, "button.btnNuevaConsulta")
            boton_volver.click()
            print("✅ Clic realizado en el botón 'Volver'")
        except (NoSuchElementException, ElementClickInterceptedException) as e:
            print(f"⚠️ Botón no disponible ({e}), intentando con el link oculto...")
            try:
                link_volver = driver.find_element(By.ID, "aNuevaConsulta")
                driver.execute_script("arguments[0].click();", link_volver)
                print("✅ Clic realizado en el link oculto 'Volver'")
            except Exception as e2:
                print(f"❌ No se pudo hacer click ni en el botón ni en el link: {e2}")
                raise

    except Exception as e:
        print(f"❌ Error Procesando SUNAT para el cliente: {cliente}, Revisar: {e}")
    finally:
        driver.refresh()
        return cod_pri,act_prin,cod_sec_1 ,act_sec_1,cod_sec_2,act_sec_2

if __name__ == "__main__":
    main()