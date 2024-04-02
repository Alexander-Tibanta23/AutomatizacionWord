import locale
from docxtpl import DocxTemplate
from datetime import datetime
import pandas as pd

locale.setlocale(locale.LC_TIME, 'es_ES')
ahora = datetime.now()

doc = DocxTemplate("at-plantilla-Documento1.docx")

nombre_juez = "Mgs. Endara Izquierdo Pablo Emilio"
dia_actual = ahora.strftime("%d")
mes_actual = ahora.strftime("%B")
año_actual = ahora.strftime("%Y")
hora_actual = ahora.strftime("%H")
minuto_actual = ahora.strftime("%M")
nombre = "ALEXANDER FRANCISCO TIBANTA"
numero_ruc = "1704472313001"
numero_cedula = "1728220441"
nombre_abogado = "Dr. Christian Santiago Izurieta Cruz"

context = {
    'nombre_juez': nombre_juez,
    'dia_actual': dia_actual,
    'mes_actual': mes_actual,
    'año_actual': año_actual,
    'hora_actual': hora_actual,
    'minuto_actual': minuto_actual,
    'nombre' : nombre,
    'numero_ruc': numero_ruc,
    'numero_cedula': numero_cedula,
    'nombre_abogado': nombre_abogado
}

doc.render(context)
doc.save("Documento-Generado.docx")