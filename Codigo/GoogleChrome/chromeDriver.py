from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from GoogleChrome.fecha_y_hora import get_fecha_actual,get_dia,get_mes, get_timestamp

import traceback
import pandas as pd
import os

#------Carpetas de Descargas y Volumen del Docker----------
carpeta_descargas = "Downloads"
ruta_carpeta_descargas = f"/app/{carpeta_descargas}"

# --- Construir ruta de Downloads por defecto ---
base_dir = os.path.dirname(os.path.abspath(__file__))
ruta_carpeta_downloads = os.path.join(base_dir, "Downloads")

def abrirDriver():
    
    #-----Configuración de Chrome para Selenium -----
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--incognito")
    #chrome_options.add_argument("--headless=new")          #"--headless" headless=new
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument('--no-sandbox')             # Lo agregue para evitar el error de sandbox
    chrome_options.add_argument('--disable-popup-blocking') # Desactivar bloqueador de ventanas emergentes
    chrome_options.add_argument("--window-size=1920,1080")  # 👈 esto es clave
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    
    # Configuracion de descargas y preferencias
    prefs = {
        "download.default_directory": ruta_carpeta_descargas,#carpeta_compañia
        "download.prompt_for_download": False,              # Evita el popup de "Guardar como"
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,         # fuerza descarga en lugar de abrir el PDF en el navegador
        "safebrowsing.enabled": True,
        "profile.default_content_settings.popups": 0,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "download.extensions_to_open": ""
        }
    chrome_options.add_experimental_option("prefs", prefs)
    #-----------------------------------
    try:
        print("🟡 Iniciando ChromeDriver con webdriver_manager")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        print("🟢 ChromeDriver iniciado correctamente")

    except Exception as e:
        print(f"❌ Error al iniciar ChromeDriver: {e}")
        traceback.print_exc()

    # Espera hasta que cargue el driver
    wait = WebDriverWait(driver, 60)
    return driver, wait

def crearCarpetas(todos_los_datos,nombre_carpeta_compañia,tipo,cia_a_verificar):
    
    if tipo == 0 :

        # Se crea carpetas dentro de Downloads,Ejemplo --> :/app/Downloads/Facturas_enviadas
        carpeta_principal = os.path.join(ruta_carpeta_descargas, nombre_carpeta_compañia)
        # Crear todas las carpetas necesarias y que existan
        print(f"📁 Creando carpeta Principal en: {ruta_carpeta_descargas}")
        os.makedirs(carpeta_principal, exist_ok=True) # <-- Esto crea la carpeta si no existe

        nombre_salida = f"Facturas_{get_timestamp()}.xlsx"
        ruta_salida_facturas = os.path.join(carpeta_principal,nombre_salida)

        nombre_log = f"EvidenciaFacturasEnviadas_{get_timestamp()}.txt"
        log_path = os.path.join(carpeta_principal, nombre_log)

        return ruta_salida_facturas , log_path

    elif tipo == 5:

        # Se crea carpetas dentro de Downloads,Ejemplo --> :/app/Downloads/Actividades_Clientes
        carpeta_principal = os.path.join(ruta_carpeta_descargas, nombre_carpeta_compañia)
        # Crear todas las carpetas necesarias y que existan
        print(f"📁 Creando carpeta Principal en: {ruta_carpeta_descargas}")
        os.makedirs(carpeta_principal, exist_ok=True) # <-- Esto crea la carpeta si no existe

        nombre_salida = f"Actividades_Clientes_Birlik_{get_dia()}_{get_mes()}.xlsx"
        ruta_salida_facturas = os.path.join(carpeta_principal,nombre_salida)

        nombre_log = "logActividad.txt"
        log_path = os.path.join(carpeta_principal, nombre_log)

        return ruta_salida_facturas , log_path ,carpeta_principal

    elif tipo == 3:

        # Se crea carpetas dentro de Downloads,Ejemplo --> :/app/Downloads/Verificacion_Cuotas_dia_mes
        ruta_carpeta_principal = os.path.join(ruta_carpeta_descargas, nombre_carpeta_compañia)
        # Crear todas las carpetas necesarias y que existan
        print(f"📁 Creando carpeta Principal en: {ruta_carpeta_descargas}")
        os.makedirs(ruta_carpeta_principal, exist_ok=True) # <-- Esto crea la carpeta si no existe

        ruta_sub_carpeta = os.path.join(ruta_carpeta_principal, cia_a_verificar)
        print(f"📁 Creando Sub Carpeta en: {ruta_carpeta_principal}")
        os.makedirs(ruta_sub_carpeta, exist_ok=True)

        nombre_log = "logVef_Cuotas.txt"
        log_path = os.path.join(ruta_sub_carpeta, nombre_log)

        return ruta_sub_carpeta , log_path

    elif tipo == 1:

        carpeta_principal = os.path.join(ruta_carpeta_descargas, nombre_carpeta_compañia)

        print(f"📁 Creando carpeta Principal en: {ruta_carpeta_descargas}")
        os.makedirs(carpeta_principal, exist_ok=True)

        nombre_salida = f"CoutasVencidas_{get_dia()}_{get_mes()}.xlsx"
        ruta_salida_cobranzas = os.path.join(carpeta_principal,nombre_salida)

        nombre_log = "Evidencia.txt"
        log_path = os.path.join(carpeta_principal, nombre_log)

        return carpeta_principal,ruta_salida_cobranzas , log_path

    else:

        #--- Esto solo se crea para cuando se cancela las cuotas

        #------ Carpeta Principal -------
        nom_carp_principal= f"Reporte_Cuotas_Diarias_{get_fecha_actual()}" #_{timestamp}
        ruta_car_principal= f"{ruta_carpeta_descargas}/{nom_carp_principal}"

        #--------Sub Carpetas y rutas --------------
        nom_carpeta_facturas ="Facturas_Descargadas"
        ruta_carpeta_facturas = f"{ruta_car_principal}/{nombre_carpeta_compañia}/{nom_carpeta_facturas}"
        nom_carpeta_comprobante = "Comprobantes_SUNAT"
        ruta_carpeta_comprobante = f"{ruta_car_principal}/{nombre_carpeta_compañia}/{nom_carpeta_comprobante}"
        nom_carpeta_errores = "Capturas_Errores"
        ruta_carpeta_errores = f"{ruta_car_principal}/{nombre_carpeta_compañia}/{nom_carpeta_errores}"
        nom_carpeta_excel = "Excel_API"
    
        # Se crea carpetas dentro de Downloads,Ejemplo --> :/app/Downloads/Reporte_Cuotas_Diarias_2025-07-21
        carpeta_principal = os.path.join(ruta_carpeta_descargas, nom_carp_principal)
    
        # Se crea carpetas dentro de Downloads, Ejemplo --> :/app/Downloads/Reporte_Cuotas_Diarias_2025-07-21/Sanitas_SCTR
        carpeta_compañia = os.path.join(ruta_car_principal, nombre_carpeta_compañia)

        # Ejemplo --> :/app/Downloads/Reporte_Cuotas_Diarias_2025-07-21/Sanitas_SCTR/Excel_API
        subcarpeta_excel = os.path.join(carpeta_compañia, nom_carpeta_excel)

        # Se crea subcarpetas dentro de la carpeta Diaria en Downloads
        subcarpeta_facturas = os.path.join(carpeta_compañia, nom_carpeta_facturas) #--> :/app/Downloads/Reporte_Cuotas_Diarias_2025-07-21/Sanitas_SCTR/Facturas_Descargadas
        subcarpeta_comprobantes = os.path.join(carpeta_compañia, nom_carpeta_comprobante) #--> :/app/Downloads/Reporte_Cuotas_Diarias_2025-07-21/Sanitas_SCTR/Comprobantes_SUNAT
        subcarpeta_errores = os.path.join(carpeta_compañia, nom_carpeta_errores) #--> :/app/Downloads/Reporte_Cuotas_Diarias_2025-07-21/Sanitas_SCTR/Capturas_Errores

        # Crear todas las carpetas necesarias y que existan
        print(f"\n📁 Creando carpeta Principal en: {ruta_carpeta_descargas}")
        os.makedirs(carpeta_principal, exist_ok=True) # <-- Esto crea la carpeta si no existe
        print("✅ Se creó correctamente:", carpeta_principal)

        print(f"\n📁 Creando carpeta de la Compañia en: {carpeta_principal}")
        os.makedirs(carpeta_compañia, exist_ok=True)
        print("✅ Se creó correctamente:", carpeta_compañia)

        print(f"\n📁 Creando sub carpetas en: {carpeta_compañia}")
        os.makedirs(subcarpeta_facturas, exist_ok=True)
        os.makedirs(subcarpeta_comprobantes, exist_ok=True)
        os.makedirs(subcarpeta_excel, exist_ok=True)
        os.makedirs(subcarpeta_errores, exist_ok=True)

        nombre_salida = f"Resultado_{nom_carp_principal}.xlsx"
        ruta_salida = os.path.join(carpeta_compañia, nombre_salida)

        salida_excel = f"Cuotas_API_{get_dia()}_{get_mes()}.xlsx"
        ruta_salida_API = os.path.join(subcarpeta_excel, salida_excel)

        salida_reporte_final= 'Reporte_Final_Cuotas.xlsx'
        ruta_maestro = os.path.join(carpeta_principal, salida_reporte_final) #--> :/app/Downloads/Reporte_Cuotas_Diarias_2025-07-21/Reporte_Final_Cuotas.xlsx

        # --- Armando el nombre del log con la MISMA base que el Excel ---
        nombre_log = nombre_salida.replace(".xlsx", ".txt")
        log_path = os.path.join(carpeta_compañia, nombre_log)

        df_final = pd.concat(todos_los_datos, ignore_index=True)
        df_final.to_excel(ruta_salida_API, index=False)
        print(f"✅ Datos del API guardados en: {ruta_salida_API}")

        return log_path,ruta_salida_API,ruta_salida,ruta_maestro,nombre_carpeta_compañia, ruta_carpeta_facturas, ruta_carpeta_comprobante, ruta_carpeta_errores,carpeta_compañia,carpeta_principal