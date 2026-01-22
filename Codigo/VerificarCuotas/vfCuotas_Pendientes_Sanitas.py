# --- Imports ---
import pandas as pd
import os
import time
import re
import sys
# --- From ----
from Apis.api_birlik import obtener_datos_cuota
from VerificarCuotas.enviarReporte import enviarReporteVerificación
from Birlik.cancelar_cuotas import fecha_creacion_birlik
from datetime import datetime
from Correo.correo_it import Tee
from GoogleChrome.fecha_y_hora import saludo_por_hora

def procesar_estado_cuenta(row):

    proforma_raw = row['PROFORMA / DOC REFERENCIA']
    fecha_emision_raw = row['FECHA COMPROBANTE']
    importe_excel = float(row['DEUDA'])

    # Validar proforma
    if pd.isna(proforma_raw) or str(proforma_raw).strip() == "":
        print(f"La Fila no tiene un código de Proforma para buscarlo en Birlik.")
        existe_resultado = "No se sabe"
        estado_resultado = "No se sabe"
        importe_resultado = "No se sabe"
        return existe_resultado,importe_resultado, estado_resultado

    # Si no está vacío, limpiar
    proforma_excel = str(proforma_raw).strip()    

    # --- Fecha ---
    if isinstance(fecha_emision_raw, (pd.Timestamp, datetime)):
        fecha_emision_dt = fecha_emision_raw.to_pydatetime()  # convertir a datetime nativo de Python
    else:
        # Si viene como string, parseamos
        fecha_emision_dt = datetime.strptime(str(fecha_emision_raw).split(" ")[0], "%Y-%m-%d")

    codigo = extraer_codigo(proforma_excel)
    codigo_extraido = extraer_valor_a_partir_tercer_indice(codigo)
    print("-------------------------")
    print(f"Fila: {proforma_excel}, Código Cuota: {codigo_extraido}, "f"Fecha Emision: {fecha_emision_dt.strftime('%d-%m-%Y')}, Importe: {importe_excel}")
                
    if fecha_emision_dt >= fecha_creacion_birlik:
            print("✅ Fecha válida. Procediendo...")

            #metodo para saber si el codigo de cuota a traves del api si existe en birlik
            datos = obtener_datos_cuota(codigo_extraido)

            if datos is None:
                print(f"No se encontro en Birlik el codigo de la cuota {codigo_extraido} y tiene estado 'Pendiente' en la compañia.")
                existe_resultado = "No"
                estado_resultado = "No se sabe"
                importe_resultado = "No se sabe"
            else:
                importe_birlik = float(datos['importe'])
                estado_birlik  = datos['estado']

                if importe_excel == importe_birlik:
                    print(f"✅ Los Importes Coinciden")
                    importe_resultado = "Coinciden"
                else:
                    print(f"❌ Los Importes No Coinciden")
                    importe_resultado = f"No Coinciden - B: {importe_birlik}"

                estado_resultado =  estado_birlik
                existe_resultado = "Si"

    else:
        print(f"⚠️ La fecha de emisión del Documento {proforma_excel} es anterior a {fecha_creacion_birlik}. No se puede procesar.")
        existe_resultado = "No figura en Birlik por la fecha de Emisión"
        estado_resultado = "No figura en Birlik por la fecha de Emisión"
        importe_resultado = "No figura en Birlik por la fecha de Emisión"

    return existe_resultado,importe_resultado,estado_resultado

def extraer_codigo(documento):
    match = re.search(r'-([0-9]+)(?:/|$)', documento)
    if match:
        return match.group(1)
    return None

def extraer_valor_a_partir_tercer_indice(documento):
    if len(documento) > 2:
        return documento[2:]  # Devuelve la cadena desde el cuarto carácter (índice 3)
    return None
  
def main(lista_descargas,ruta_salida_vc,log_path):
    
    # --- Redireccionar todo el stdout al archivo de log ---
    original_stdout = sys.stdout  # Guarda referencia original a la consola
    with open(log_path, "w", encoding="utf-8") as log_file:
        #sys.stdout = log_file
        sys.stdout = Tee(sys.stdout, log_file)

        try:
            cia = "Sanitas"
            # Aquí va TODO tu código que imprime cosas
            print(f"📁 Iniciando Verificación de cuotas para {cia}...")

            try:
                ruta_excel_descargado = lista_descargas[0]   # o [-1]
                df = pd.read_excel(ruta_excel_descargado, engine="openpyxl",skiprows=5)
            except Exception as e:
                print(f"❌ Error al leer el archivo Excel: {e}")
                sys.exit() #return 

            total_filas = len(df)

            # Inicializa columnas si no existen
            df["Existe en Birlik"] = ""
            df["Importe en Birlik"] = ""
            df["Estado en Birlik"] = ""

            # 2. Iterar sobre cada fila
            for index, row in df.iterrows():
                print(f"\n--- Procesando fila {index + 2} de {total_filas + 1} ---")

                nom_excel_vfCuotas = "Cuotas_Pendientes_Sanitas.xlsx"
                ruta_resultado = os.path.join(ruta_salida_vc,nom_excel_vfCuotas)

                try:
                    existe_birlik,importe_birlik , estado_birlik = procesar_estado_cuenta(row)

                    df.at[index, "Existe en Birlik"] = existe_birlik
                    df.at[index, "Importe en Birlik"] = importe_birlik
                    df.at[index, "Estado en Birlik"] = estado_birlik

                    # 💾 Guardar luego de procesar cada fila
                    df.to_excel(ruta_resultado, index=False)

                    print(f"✔ Fila {index} guardada correctamente.")
                except Exception as e:
                    print(f"❌ Error en fila {index},Detalle: {e}")
                    df.to_excel(ruta_resultado, index=False)
                    continue
                
                time.sleep(3)

            print(f"\n✅ Proceso finalizado. Todo el log ha sido guardado en:\n{log_path}")

            saludo = saludo_por_hora()

            enviarReporteVerificación(saludo,cia,ruta_resultado)

        finally:
            sys.stdout = original_stdout  # Restaura la salida estándar (la consola)

if __name__ == "__main__":
    main()
