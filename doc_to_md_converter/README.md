# Doc to Markdown Converter

Convertidor de documentos Word (.doc, .docx) a Markdown con una interfaz web moderna y con visualizador de la covercion tipo html.

## Requisitos

- Python 3.7+
- pip
- Entorno virtual (venv)

## Instalación y Ejecución

### Opción 1: PowerShell (Windows)

```powershell
# 1. Navegar al directorio del proyecto
cd .\doc_to_md_converter

# 2. Crear entorno vitrual
python -m venv venv

# 3. Activar el entorno virtual
.\venv\Scripts\Activate.ps1

# 4. Instalar dependencias
pip install flask python-docx mammoth

# 5. Ejecutar la aplicación
python app.py
```

### Opción 2: Command Prompt (CMD) - Windows

```cmd
# 1. Navegar al directorio del proyecto
cd c:\Users\2905038\doc_to_md_converter

# 2. Activar el entorno virtual
venv\Scripts\activate.bat

# 3. Instalar dependencias
pip install flask python-docx mammoth

# 4. Ejecutar la aplicación
python app.py
```

### Opción 3: Linux/Mac

```bash
# 1. Navegar al directorio del proyecto
cd /path/to/doc_to_md_converter

# 2. Activar el entorno virtual
source venv/bin/activate

# 3. Instalar dependencias
pip install flask python-docx mammoth

# 4. Ejecutar la aplicación
python app.py
```

## Accesso a la Aplicación

Una vez ejecutado `python app.py`, la aplicación estará disponible en:

```
http://localhost:5000
```

Abre tu navegador y accede a esa dirección para usar la interfaz web.

## Características

- ✅ Conversión de archivos .doc y .docx a Markdown
- ✅ Interfaz web intuitiva
- ✅ Descarga de archivos convertidos
- ✅ Soporte para múltiples archivos
- ✅ Límite de 50 MB por archivo

## Estructura del Proyecto

```
doc_to_md_converter/
├── app.py                 # Aplicación Flask principal
├── doc_to_md.py           # Lógica de conversión
├── devkit.litcoffee       # Configuración de desarrollo
├── templates/
│   └── index.html         # Interfaz web
├── uploads/               # Carpeta para archivos subidos
├── outputs/               # Carpeta para archivos convertidos
├── venv/                  # Entorno virtual
└── README.md              # Este archivo
```

## Dependencias Principales

- **Flask**: Framework web para Python
- **python-docx**: Lectura de archivos Word (.docx)
- **mammoth**: Conversión HTML → Markdown

## Notas

- Los archivos subidos se almacenan temporalmente en la carpeta `uploads/`
- Los archivos convertidos se guardan en `outputs/`
- El tamaño máximo de archivo es de 50 MB
- Se soportan formatos `.doc` y `.docx`

## Troubleshooting

### Error: "No module named 'flask'"
```powershell
pip install flask
```

### Error: "No module named 'docx'"
```powershell
pip install python-docx
```

### Error: "No module named 'mammoth'"
```powershell
pip install mammoth
```

### La aplicación no inicia
Verifica que el entorno virtual esté activado (deberías ver `(venv)` en la terminal)

## Autor

Convertidor Doc to Markdown - 2026
