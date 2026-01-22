# -- Froms --
from Correo.correo_it import Tee
from Apis.api_birlik import obtener_datos_cuota
from Excels.estilosExcel import guardar_excel_con_formato_solo_ajustar_columnas
from VerificarCuotas.enviarReporte import enviarReporteVerificación
from GoogleChrome.fecha_y_hora import saludo_por_hora
# -- Imports --
import os
import time
import pandas as pd
import sys

def procesar_estado_cuenta(row):
 
    columna_codigo = "Número" if "Número" in row else ("PROFORMA" if "PROFORMA" in row else "Codigo Cuota")
    codigoCuota = str(row[columna_codigo]).strip()

    columna_monto = "PRIMA" if "PRIMA" in row else ("MONTO" if "MONTO" in row else "Ctas.  por Cobrar")
    monto_raw = row[columna_monto]
    monto = float(str(monto_raw).replace(",", "."))
    
    print("-------------------------")
    print(f"Fila: Código Cuota: {codigoCuota}, Importe: {monto}")
                
    #metodo para saber si el codigo de cuota a traves del api si existe en birlik
    datos = obtener_datos_cuota(codigoCuota)

    if datos is None:
        print("Tiene estado 'Pendiente' en la compañía.")
        existe_resultado = "No"
        estado_resultado = "No se sabe"
        importe_resultado = "No se sabe"
        contexto = "No existe la proforma en Birlik."
        usuario_birlik = "N/A"
    else:
        importe_birlik = float(datos['importe'])
        estado_birlik  = datos['estado']
        usuario_birlik  = datos['fkusuario']

        # Validación de importes con tolerancia
        diferencia = abs(monto - importe_birlik)


        if diferencia <= 0.05: # monto == importe_birlik
            print(f"✅ Los Importes Coinciden")
            importe_resultado = "Coinciden"
            contexto = "Registrado en Birlik con importe iguales."
        else:
            print(f"❌ Los Importes No Coinciden")
            importe_resultado = f"No Coinciden - B: {importe_birlik}"
            contexto = "Registrado en Birlik con importes diferentes"

        estado_resultado =  estado_birlik
        existe_resultado = "Si"

    return existe_resultado,importe_resultado,estado_resultado,usuario_birlik,contexto,

def main(lista_descargas,ruta_salida_vc,log_path):

    original_stdout = sys.stdout
    with open(log_path, "w", encoding="utf-8") as log_file:
        sys.stdout = Tee(sys.stdout, log_file)
        try:
            cia = "La Positiva"
            print(f"📁 Iniciando Verificación de cuotas para {cia}...")
            try:
                ruta_excel_descargado = lista_descargas[0]
                df = pd.read_excel(ruta_excel_descargado, engine="openpyxl",skiprows=6)
            except Exception as e:
                print(f"❌ Error al leer el archivo Excel: {e}")
                sys.exit()
            
            total_filas = len(df)

            df["Existe en Birlik"] = ""
            df["Importe en Birlik"] = ""
            df["Estado en Birlik"] = ""
            df["Responsable"] = ""
            df["Contexto"] = ""

            for index, row in df.iterrows():
                print(f"\n--- Procesando fila {index + 2} de {total_filas + 1} ---")

                nom_excel_vfCuotas = "Cuotas_Pendiente_Positiva.xlsx"
                ruta_resultado = os.path.join(ruta_salida_vc,nom_excel_vfCuotas)

                try:
                    existe_birlik,importe_birlik , estado_birlik ,usuario_birlik,contexto= procesar_estado_cuenta(row)

                    df.at[index, "Existe en Birlik"] = existe_birlik
                    df.at[index, "Importe en Birlik"] = importe_birlik
                    df.at[index, "Estado en Birlik"] = estado_birlik
                    df.at[index, "Responsable"] = usuario_birlik
                    df.at[index, "Contexto"] = contexto

                    df.to_excel(ruta_resultado, index=False)
                    print(f"✔ Fila {index} guardada correctamente.")
                except Exception as e:
                    print(f"❌ Error en fila {index},Detalle: {e}")
                    df.to_excel(ruta_resultado, index=False)
                    continue

                time.sleep(3)

            print(f"\n✅ Proceso finalizado. Todo el log ha sido guardado en:\n{log_path}")

            guardar_excel_con_formato_solo_ajustar_columnas(ruta_resultado,'Sheet1')

            saludo = saludo_por_hora()

            enviarReporteVerificación(saludo,cia,ruta_resultado)

        finally:
            sys.stdout = original_stdout  # Restaura la salida estándar (la consola)

if __name__ == "__main__":
    main()