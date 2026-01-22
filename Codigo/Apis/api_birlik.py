import pandas as pd
import requests
import os

from openpyxl import load_workbook
from requests.exceptions import ReadTimeout, ConnectTimeout, RequestException

#----CREDENCIALES API BIRLIK -------
API_KEY = os.getenv("API_KEY")
AFTER_API_KEY = os.getenv("AFTER_API_KEY")
API_KEY_HEADER = AFTER_API_KEY+" "+API_KEY

def guardarDatosAPI_excel(datos_cobranza,ruta_salida_API):
    excel_json = pd.DataFrame(datos_cobranza)
    excel_json.to_excel(ruta_salida_API, index=False)

    wb = load_workbook(ruta_salida_API)
    ws = wb.active

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter  # Ej: "A", "B", etc.
        for cell in col:
           try:
              if cell.value:
                max_length = max(max_length, len(str(cell.value)))
           except:
                pass
        adjusted_width = max_length + 2
        ws.column_dimensions[col_letter].width = adjusted_width

    wb.save(ruta_salida_API)

# Obtener todos los datos de una cuota por su codigo
def obtener_datos_cuota(codigo_cuota):
    url = f"https://plataformabirlik.azurewebsites.net/api/cuotaapi/codigo/{codigo_cuota}"
    headers = {
        "Authorization": API_KEY_HEADER
    }
    try:
        response = requests.get(url, headers=headers, timeout=5)
        if response.status_code == 200:
            data = response.json()
            return {
                "estado": data.get("estado"),
                "importe": data.get("importe"),
                "fkusuario" : data.get("fkusuario")
                }
        else:
            print(f"[API]: {response.text}")
            return None
    except Exception as e:
        print(f"[API] Error consultando la API para el codigo de  cuota {codigo_cuota}: {e}")
        return None

# Obtener solo estado de esa cuota por Id
def obtener_estado_cuota(id_cuota):
    url = f"https://plataformabirlik.azurewebsites.net/api/cuotaapi/{id_cuota}/estado"
    headers = {
        "Authorization": API_KEY_HEADER
    }
    try:
        response = requests.get(url, headers=headers, timeout=5)
        if response.status_code == 200:
            data = response.json()
            return data.get("estado", None)
        else:
            print(f"[API] No se pudo obtener el estado de la cuota {id_cuota}: {response.text}")
            return None
    except Exception as e:
        print(f"[API] Error consultando la API para cuota {id_cuota}: {e}")
        return None

def main():
    mensaje = "Hola"
    print(mensaje)

if __name__ == "__main__":
    main()

def consultarAPI(url,ids_compania):

    # Lista para acumular resultados
    todos_los_datos = []

    #print(f"📡 Consultando API de Birlik para todas las CIAs.")
    for fk_compania in ids_compania:
        #print(f"📡 Consultando API para compañía {fk_compania}...")

        #-- Esta API para enviar factura
        datos_cobranza = ObtenerListadeDatosporFk_Compania(url,fk_compania)
    
        if datos_cobranza is None:
            #print(f"❌ Error al obtener datos para compañía {fk_compania}.")
            continue

        if not datos_cobranza:  # Si es [] o vacío
            #print(f"⚠️ No hay cuotas para la compañía {fk_compania}.")
            continue

        # Si hay datos, convertir a DataFrame y acumular
        df = pd.DataFrame(datos_cobranza)
        if df.empty:
            print(f"⚠️ DataFrame vacío para la compañía {fk_compania}.")
            continue

        todos_los_datos.append(df)
        #print(f"✅ Datos de la Cuota extraído correctamente de la compañía {fk_compania}")

    return todos_los_datos

def guardar_observacion_cuota(id_cuota, observacion):
    url = "https://plataformabirlik.azurewebsites.net/api/cuota/guardar-observacion"
    headers = {
        "Authorization": API_KEY_HEADER,  # si tu API necesita auth
        "Content-Type": "application/json"
    }
    payload = {
        "IdCuota": id_cuota,
        "Observacion": observacion
    }

    try:
        response = requests.post(url, headers=headers, json=payload, timeout=5)
        if response.status_code == 200:
            data = response.json()
            print(f"[API] Observación guardada en cuota {id_cuota}: {data}")
            return True
        else:
            print(f"[API] Error al guardar observación en cuota {id_cuota}: {response.text}")
            return False
    except Exception as e:
        print(f"[API] Excepción al llamar API para cuota {id_cuota}: {e}")
        return False

# Metodo para obtener una lista de datos por fk_compania
def ObtenerListadeDatosporFk_Compania(url,fk_compania):
    urlfinal = f"{url}{fk_compania}"
    headers = {
        "Authorization": API_KEY_HEADER
    }
    # try:
    #     response = requests.get(urlfinal, headers=headers, verify=True, timeout=(10, 60)) #False
    #     if response.status_code == 200:
    #         datos = response.json()
    #         return datos  # Es una lista de diccionarios
    #     else:
    #         print(f"[API] Error al obtener datos de la Cuota: {response.status_code} - {response.text}")
    #         return None
    # except Exception as e:
    #     print(f"[API] Error conectando con la API: {e}")
    #     return None
    try:
        response = requests.get(urlfinal,headers=headers,verify=True,timeout=(10, 60))
        if response.status_code == 200:
            return response.json()  # lista de dicts
        print(f"[API] Error HTTP {response.status_code}: {response.text}")
        return None
    except ReadTimeout:
        print("[API] ⏱️ Timeout leyendo respuesta de la API")
        return None
    except ConnectTimeout:
        print("[API] 🔌 Timeout conectando a la API")
        return None
    except RequestException as e:
        print(f"[API] ❌ Error de requests: {e}")
        return None