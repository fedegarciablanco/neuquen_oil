# Neuquén Oil

Base reproducible para analizar datos del sector petrolero de Neuquén.

## Primer uso en VS Code

1. Instalar Python 3.12 para Windows, con el lanzador `py` habilitado.
2. Abrir esta carpeta como espacio de trabajo en VS Code.
3. Instalar las extensiones recomendadas cuando VS Code las sugiera.
4. Crear el entorno con `py -3.12 -m venv .venv`.
5. Activarlo con `.venv\\Scripts\\Activate.ps1` e instalar las dependencias con
   `python -m pip install -r requirements.txt`.
6. Seleccionar el intérprete `Python 3.12 (.venv)` si VS Code no lo detecta automáticamente.

El entorno `.venv` no se guarda en Git.

## Calidad

Desde VS Code se pueden ejecutar las tareas **Calidad: revisar** y **Pruebas:
ejecutar**. Al guardar archivos Python, Ruff formatea el código y ordena imports.
