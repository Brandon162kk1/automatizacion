import token
import vfCuotas_Pendientes_Positiva
import vfCuotas_Pendientes_Sanitas
import os
import requests

from Correo.correo_it import guardar_excel,EMAIL_ACCOUNT
from GoogleChrome.chromeDriver import crearCarpetas,get_fecha_actual

asunto = os.getenv("asunto")
token = os.getenv("token")
message_id = os.getenv("message_id")

def main():

        try:

            rutas_excel = []
            cia_a_verificar = ""
            log_path = None
            ruta_salida_vc =  None
            headers = {
                'Authorization': 'Bearer ' + token,
                'Content-Type': 'application/json'
            }
    
            # Definir "mapa" de prefijo -> compañía 
            prefijos = {
                "VC_SANI": "Sanitas",
                "VC_POSI": "Positiva",
                "VC_PACI": "Pacífico",
                "VC_RIMAC": "Rímac"
            }
            
            # Normalizar asunto
            asunto_limpio = asunto.strip().upper()
            # Buscar qué prefijo aplica
            cia_a_verificar = None
            for prefijo, compania in prefijos.items():
                if prefijo.upper() in asunto_limpio:
                    cia_a_verificar = compania
                    break

            if cia_a_verificar:
                try:
                    # Obtener adjuntos
                    attach_url = f"https://graph.microsoft.com/v1.0/users/{EMAIL_ACCOUNT}/messages/{message_id}/attachments"
                    attach_resp = requests.get(attach_url, headers=headers)

                    if attach_resp.status_code == 200:
                        attachments = attach_resp.json().get('value', [])
                        for attach in attachments:
                            if attach['@odata.type'] == "#microsoft.graph.fileAttachment":
                                nombre_archivo = attach['name']

                                if nombre_archivo.lower().endswith((".xls", ".xlsx")):

                                    json_vacio = {}
                                    nom_car_pri= f"Verificacion_Cuotas_{get_fecha_actual()}"
                                    tipo = 3
                                    ruta_salida_vc,log_path = crearCarpetas(json_vacio,nom_car_pri,tipo,cia_a_verificar)

                                    ruta_guardada = guardar_excel(ruta_salida_vc, nombre_archivo, attach['contentBytes'])
                                    rutas_excel.append(ruta_guardada)
                                else:
                                    raise Exception(f"📂 Archivo ignorado (no es Excel): {nombre_archivo}")
                    else:
                        raise Exception(f"Error al obtener adjuntos: {attach_resp.status_code}, {attach_resp.text}")

                except Exception as e:
                    raise Exception(f"Error al procesar correo ({cia_a_verificar}): {e}")
            else:
                raise Exception(f"Asunto no reconocido: {asunto}")

            if asunto == "VC_POSI":
                vfCuotas_Pendientes_Positiva.main(rutas_excel,ruta_salida_vc,log_path)
            else:
                vfCuotas_Pendientes_Sanitas.main(rutas_excel,ruta_salida_vc,log_path)
            #---------------------------------------------------------------------------------------------------------
        except Exception as e:
            print(f"❌ Error inesperado: {e}")
        finally:
            pass

if __name__ == "__main__":
    main()