import locale
from docxtpl import DocxTemplate
from datetime import datetime
import pandas as pd

locale.setlocale(locale.LC_TIME, 'es_ES')
ahora = datetime.now()

doc = DocxTemplate("at-plantilla-Documento1.docx")

dia_actual = ahora.strftime("%d")
mes_actual = ahora.strftime("%B")
año_actual = ahora.strftime("%Y")
hora_actual = ahora.strftime("%H")
minuto_actual = ahora.strftime("%M")

my_context = {
    'dia_actual': dia_actual,
    'mes_actual': mes_actual,
    'año_actual': año_actual,
    'hora_actual': hora_actual,
    'minuto_actual': minuto_actual,
}

df = pd.read_csv('fake-data.csv', sep=';')

for index, fila in df.iterrows():
    context = {
        'nombre_juez': fila['Juez'],
        'nombre' : fila['nombre'],
        'numero_ruc': fila['ruc'],
        'numero_cedula': fila['cedula'],
        'nombre_abogado': fila['abogado'],
    }

    context.update(my_context)
    
    doc.render(context)
    doc.save(f"Documento-Generado_{index}.docx")