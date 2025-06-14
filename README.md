# AutomatizacionWord

Creación de una aplicación de escritorio para la automatización de documentos Word usando plantillas y datos desde archivos Excel.

## Requisitos
- Python 3.8 o superior
- pip (gestor de paquetes de Python)

## Instalación de dependencias
Desde la raíz del proyecto, ejecuta:

```powershell
pip install -r requirements.txt
```

## Ejecución
Para iniciar la aplicación para personas naturales:

```powershell
python .\model\ProvidenciaNatural.py
```

Para personas jurídicas:

```powershell
python .\model\ProvidenciaJurica.py
```

## Estructura del proyecto
- `model/` — Scripts principales de automatización (no se suben a git)
- `data/` — Bases de datos en Excel
- `templates/` — Plantillas Word
- `output/` — Documentos generados y ejecutables
- `styles/` — Archivos de estilos para la interfaz

## Ejemplo de uso
1. Ejecuta el script correspondiente.
2. Selecciona un registro de la base de datos.
3. Genera el documento Word usando la plantilla.
4. El documento se guardará en la carpeta `output/`.

## Notas
- Las carpetas `model`, `JURIDICA`, `NATURAL` y `output` están excluidas del control de versiones por `.gitignore`.
- Si falta alguna dependencia, instálala con `pip install nombre_paquete`.