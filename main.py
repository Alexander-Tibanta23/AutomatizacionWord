import tkinter as tk
from tkinter import ttk
from docxtpl import DocxTemplate
from datetime import datetime, timedelta

# Variables globales
document_data = []

# Función para insertar datos desde la GUI a la lista y actualizar el Treeview
def insert_row():
    nombre_juez = nombre_juez_entry.get()
    nombre = nombre_entry.get()
    numero_ruc = numero_ruc_entry.get()
    numero_cedula = numero_cedula_entry.get()
    nombre_abogado = nombre_abogado_entry.get()

    row_data = {
        'Juez': nombre_juez,
        'nombre': nombre,
        'ruc': numero_ruc,
        'cedula': numero_cedula,
        'abogado': nombre_abogado,
    }
    document_data.append(row_data)
    
    # Actualizar el Treeview
    update_treeview()

    # Limpiar campos
    nombre_juez_entry.delete(0, tk.END)
    nombre_entry.delete(0, tk.END)
    numero_ruc_entry.delete(0, tk.END)
    numero_cedula_entry.delete(0, tk.END)
    nombre_abogado_entry.delete(0, tk.END)

# Actualiza el contenido del Treeview
def update_treeview():
    for i in tree.get_children():
        tree.delete(i)
    for data in document_data:
        tree.insert('', 'end', values=(data['Juez'], data['nombre'], data['ruc'], data['cedula'], data['abogado']))

# Función para generar documentos Word con fecha y hora actualizadas
def generate_documents():
    inicio = datetime.now()
    doc = DocxTemplate("at-plantilla-Documento1.docx")
    for index, fila in enumerate(document_data):
        ahora = inicio + timedelta(minutes=index)  # Incrementa un minuto por cada documento
        my_context = {
            'dia_actual': ahora.strftime("%d"),
            'mes_actual': ahora.strftime("%B"),
            'año_actual': ahora.strftime("%Y"),
            'hora_actual': ahora.strftime("%H"),
            'minuto_actual': ahora.strftime("%M"),
        }
        combined_context = {**fila, **my_context}
        doc.render(combined_context)
        doc.save(f"Documento-Generado_{index}.docx")

# Configuración de la GUI
root = tk.Tk()
root.title("Generador de Documentos")

main_frame = ttk.Frame(root, padding="10")
main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Campos de entrada y sus etiquetas
ttk.Label(main_frame, text="Nombre del Juez:").grid(row=0, column=0, pady=2, sticky=tk.W)
nombre_juez_entry = ttk.Entry(main_frame)
nombre_juez_entry.grid(row=0, column=1, pady=2, sticky=(tk.W, tk.E))

ttk.Label(main_frame, text="Nombre:").grid(row=1, column=0, pady=2, sticky=tk.W)
nombre_entry = ttk.Entry(main_frame)
nombre_entry.grid(row=1, column=1, pady=2, sticky=(tk.W, tk.E))

ttk.Label(main_frame, text="Número RUC:").grid(row=2, column=0, pady=2, sticky=tk.W)
numero_ruc_entry = ttk.Entry(main_frame)
numero_ruc_entry.grid(row=2, column=1, pady=2, sticky=(tk.W, tk.E))

ttk.Label(main_frame, text="Número Cédula:").grid(row=3, column=0, pady=2, sticky=tk.W)
numero_cedula_entry = ttk.Entry(main_frame)
numero_cedula_entry.grid(row=3, column=1, pady=2, sticky=(tk.W, tk.E))

ttk.Label(main_frame, text="Nombre del Abogado:").grid(row=4, column=0, pady=2, sticky=tk.W)
nombre_abogado_entry = ttk.Entry(main_frame)
nombre_abogado_entry.grid(row=4, column=1, pady=2, sticky=(tk.W, tk.E))

# Botones
insert_button = ttk.Button(main_frame, text="Insertar Datos", command=insert_row)
insert_button.grid(row=5, column=0, columnspan=2, pady=5)

generate_button = ttk.Button(main_frame, text="Generar Documentos", command=generate_documents)
generate_button.grid(row=6, column=0, columnspan=2, pady=5)

# Configuración del Treeview
tree_frame = ttk.Frame(root)
tree_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

columns = ('nombre_juez', 'nombre', 'numero_ruc', 'numero_cedula', 'nombre_abogado')
tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
tree.heading('nombre_juez', text="Nombre del Juez")
tree.heading('nombre', text="Nombre")
tree.heading('numero_ruc', text="Número RUC")
tree.heading('numero_cedula', text="Número Cédula")
tree.heading('nombre_abogado', text="Nombre del Abogado")

tree.grid(row=0, column=0, sticky='nsew')

# Añadir Scrollbar
scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
scrollbar.grid(row=0, column=1, sticky='ns')
tree.configure(yscroll=scrollbar.set)

root.mainloop()