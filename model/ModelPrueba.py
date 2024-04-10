import locale
from docxtpl import DocxTemplate
from datetime import datetime, timedelta
import pandas as pd

locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')  # Asegúrate de que el código de localización sea válido en tu sistema
inicio = datetime.now()

doc = DocxTemplate("at-plantilla-Documento1.docx")

df = pd.read_csv('fake-data.csv')

for index, fila in df.iterrows():
    # Define el momento actual para este documento específico
    ahora = inicio + timedelta(minutes=index)  # Incrementa un minuto por cada documento

    # Contexto de fecha y hora actualizado
    my_context = {
        'dia_actual': ahora.strftime("%d"),
        'mes_actual': ahora.strftime("%B"),
        'año_actual': ahora.strftime("%Y"),
        'hora_actual': ahora.strftime("%H"),
        'minuto_actual': ahora.strftime("%M"),
    }

    # Contexto específico de la fila
    context = {
        'nombre_juez': fila['Juez'],
        'nombre': fila['nombre'],
        'numero_ruc': fila['ruc'],
        'numero_cedula': fila['cedula'],
        'nombre_abogado': fila['abogado'],
    }

    # Combina los contextos
    combined_context = {**context, **my_context}
    
    doc.render(combined_context)
    doc.save(f"Documento-Generado_{index}.docx")