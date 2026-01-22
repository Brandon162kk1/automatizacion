#-- Imports --
import base64
import json
import requests
import os

# --- Variables de Entorno ---
remitente = os.getenv("remitente")
client_id = os.getenv("client_id")
client_secret = os.getenv("client_secret")
tenant_id = os.getenv("TENANT_ID")
SCOPE = os.getenv("SCOPE")

def formato_correos(lista):
    return [{"emailAddress": {"address": correo}} for correo in lista]

def enviarCorreoIT(destinatarios_to,destinatarios_cc,asunto, mensaje, ruta_imagen=None,lista_adjuntos=None):

    token_endpoint = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
 
    if lista_adjuntos is None:
        lista_adjuntos = []
    
    # Paso 1: Obtener el token
    token_data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": SCOPE
    }
    token_response = requests.post(token_endpoint, data=token_data)
    access_token = token_response.json().get("access_token")
 
    if not access_token:
        print("❌ No se pudo obtener el token.")
        return
 
    # Paso 2: Construir el cuerpo del mensaje
    email_body = {
        "message": {
            "subject": asunto,
            "body": {
                "contentType": "HTML",
                "content": f"""
        <html>
        <body>
        {mensaje}
        </body>
        </html>
                """
            },
            "toRecipients": formato_correos(destinatarios_to),
            "ccRecipients": formato_correos(destinatarios_cc),
        },
        "saveToSentItems": "true"
    }
 
    attachments = []

    for archivo_path in lista_adjuntos:
        if os.path.exists(archivo_path):
            with open(archivo_path, "rb") as f:
                contenido_base64 = base64.b64encode(f.read()).decode('utf-8')
                attachments.append({
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": os.path.basename(archivo_path),
                    "contentBytes": contenido_base64,
                    "contentType": "application/octet-stream"
                })
        else:
            print(f"\n⚠️ Archivo no encontrado: {archivo_path}")

    if ruta_imagen and os.path.exists(ruta_imagen):
        with open(ruta_imagen, "rb") as f:
            contenido_base64 = base64.b64encode(f.read()).decode('utf-8')
            attachments.append({
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": os.path.basename(ruta_imagen),
                "contentBytes": contenido_base64,
                "contentType": "image/png"
            })
 
    # Solo agregar attachments si existen
    if attachments:
         email_body["message"]["attachments"] = attachments

    # Paso 3: Enviar el correo con Microsoft Graph
    url_envio = f"https://graph.microsoft.com/v1.0/users/{remitente}/sendMail"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    try:

        response = requests.post(
            url_envio,
            headers=headers,
            data=json.dumps(email_body),
            timeout=30
        )
        if response.status_code == 202:
            print(f"📧 Correo enviado correctamente a {destinatarios_to} y con copia a {destinatarios_cc}.")
        else:
            print("❌ Error al enviar correo:", response.status_code, response.text)
    except requests.exceptions.Timeout:
        print("⏰ El envío del correo tardó demasiado y fue cancelado (timeout).")
    except Exception as e:
        print(f"❌ Error inesperado al enviar el correo: {e}")

def enviarCaptcha(para, copia, puerto, cia, imagen):
    url = f"http://jishucloud.redirectme.net:{puerto}"

    asunto = f"🧩 Resolver Captcha en {cia}"
    mensaje = f"""
    <p>Ingresar al siguiente enlace y resolver el captcha manualmente si es que aparece.</p>
    <p>
        👉 <a href="{url}" target="_blank">{url}</a>
    </p>
    <p>Finaliza con clic en <b>Ingresar</b>.</p>
    """

    enviarCorreoIT(para, copia, asunto, mensaje, imagen, None)
