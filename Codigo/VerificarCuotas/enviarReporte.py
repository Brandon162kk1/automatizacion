# -- Froms --
from Correo.correo_it import enviarCorreoIT
# -- Imports --

destinatarios = ["brandon.rodriguez@jishu.com.pe"]
destinatarios_copia = ["miguel.aquino@jishu.com.pe"]

# destinatarios = ["harold.li@birlik.com.pe","yanian.chu@birlik.com.pe","neih.chu@birlik.com.pe"]
# destinatarios_copia=  ["miguel.aquino@jishu.com.pe","brandon.rodriguez@jishu.com.pe"]

def enviarReporteVerificación(saludo,cia,ruta_resultado):

    asunto = f"Verificación de cuotas en Birlik"
    mensaje_html = f"""
            <html>
              <body style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
                <p>Estimados(as),</p>
                <p>{saludo},</p>

                <p>
                  Les adjunto los resultados obtenidos de cada cuota pendiente en Birlik con respecto a <b>{cia}</b>.
                </p>

                <p>
                  Asimismo, les informamos que se ha habilitado un proceso de
                  automatización para verificar de manera más rápida y precisa las cuotas con estado de cuenta 'Pendiente' en
                  las compañías de seguros, facilitando así la disponibilidad de la información en el menor tiempo posible.
                </p>

                <p>
                  Para poder ejecutar esta automatización, es necesario enviar un correo al área de Tecnología (IT)
                  con el siguiente formato en el asunto:
                </p>

                <p><i>VC_(*4 primeros caracteres de la compañía a consultar)</i></p>

                <p><b>Ejemplos:</b></p>
                <ul>
                  <li>La Positiva → <code>VC_POSI</code></li>
                  <li>Sanitas → <code>VC_SANI</code></li>
                  <li>Pacífico → <code>VC_PACI</code></li>
                  <li>Rímac → <code>VC_RIMA</code></li>
                </ul>

                <br>
                <p>Saludos cordiales,</p>
                <p><b>Departamento Tecnológico - Birlik</b></p>
              </body>
            </html>
            """
    lista_adjuntos = [ruta_resultado]
    enviarCorreoIT(destinatarios,destinatarios_copia,asunto, mensaje_html,None,lista_adjuntos)
    #enviarCorreoIT(destinatarios,destinatarios_copia,asunto, mensaje_html, lista_adjuntos)