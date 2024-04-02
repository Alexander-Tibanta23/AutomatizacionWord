import locale
from datetime import datetime
import pandas as pd

locale.setlocale(locale.LC_TIME, 'es_ES')

ahora = datetime.now()

dia_actual = ahora.strftime("%d")
mes_actual = ahora.strftime("%B")
año_actual = ahora.strftime("%Y")
hora_actual = ahora.strftime("%H")
minuto_actual = ahora.strftime("%M")
nombre = "ALEXANDER FRANCISCO TIBANTA"
numero_ruc = "1704472313001"
numero_cedula = "1728220441"
nombre_abogado = "Dr. Christian Santiago Izurieta Cruz"

# Imprimir resultados
print(f"Día actual: {dia_actual}")
print(f"Mes actual: {mes_actual}")
print(f"Año actual: {año_actual}")
print(f"Hora actual: {hora_actual}")
print(f"Minutos actuales: {minuto_actual}")

df = pd.read_csv('fake-data.csv')
print(df.columns)