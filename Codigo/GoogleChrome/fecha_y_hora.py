from datetime import datetime
import pytz

# Definir la zona horaria de Lima
tz_peru = pytz.timezone("America/Lima")

# #-------FECHA  Y TIEMPO GLOBAL---------
def get_fecha_hoy():
    return datetime.now(tz_peru)

def get_timestamp():
    return datetime.now(tz_peru).strftime('%Y%m%d_%H%M%S')

def get_fecha_actual():
    return datetime.now(tz_peru).strftime("%Y-%m-%d")

def get_anio():
    return datetime.now(tz_peru).strftime("%Y")

def get_dia():
    return datetime.now(tz_peru).strftime("%d")

def get_mes():
    return datetime.now(tz_peru).strftime("%m")

def get_hora():
    return datetime.now(tz_peru).strftime("%H")

def get_minuto():
    return datetime.now(tz_peru).strftime("%M")

def get_segundo():
    return datetime.now(tz_peru).strftime("%S")

def get_pos_fecha_dmy():
    return datetime.now(tz_peru).strftime("%d/%m/%Y")

def saludo_por_hora():

    if 6 <= get_fecha_hoy().hour < 12:
        return "Buenos días"
    elif 12 <= get_fecha_hoy().hour < 18:
        return "Buenas tardes"
    else:
        return "Buenas noches"