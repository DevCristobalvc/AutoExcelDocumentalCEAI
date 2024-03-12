# AUTO EXCEL DOCUMENTAL CEAI

Este script de Python automatiza el proceso de generación de archivos de control en formato Excel a partir de un archivo Excel fuente, empleando una plantilla predefinida. Está especialmente diseñado para mejorar la eficiencia en la gestión documental, permitiendo la creación masiva y automatizada de documentos de control con datos específicos extraídos del archivo fuente.

## Funcionalidades

- **Lectura automática de archivos Excel**: El script lee un archivo Excel fuente, extrayendo los datos necesarios para completar la plantilla de control.
- **Generación de documentos de control**: Utiliza una plantilla Excel predefinida para crear nuevos documentos de control con los datos extraídos, almacenándolos en una ubicación específica.
- **Interfaz Gráfica de Usuario (GUI)**: Ofrece una interfaz gráfica sencilla e intuitiva para seleccionar el archivo fuente, la plantilla y el directorio de destino de los archivos de control.
- **Verificación de duplicados**: Antes de crear un nuevo archivo de control, verifica si ya existe en el destino para evitar la creación de duplicados.
- **Integración de imagen en la GUI**: Permite la inclusión de un logotipo en la interfaz gráfica, mejorando la identidad visual de la herramienta.

## Pre-requisitos

Para ejecutar este script, necesitarás Python instalado en tu sistema y las siguientes librerías:

- pandas
- openpyxl
- tkinter
- PIL (Python Imaging Library)

## Instalación de dependencias

Ejecuta el siguiente comando en tu terminal para instalar las dependencias necesarias:

```bash
pip install pandas openpyxl Pillow
```

## Instrucciones de uso

1. **Ejecuta el script**: Utiliza tu terminal o un entorno de desarrollo que soporte la ejecución de scripts de Python para ejecutar el script `.py`.

2. **Interfaz gráfica**:
   - **Selecciona la ruta del archivo fuente**: Busca y elige el archivo Excel que servirá de fuente para la generación de los archivos de control.
   - **Selecciona la ruta de la plantilla control**: Indica dónde se encuentra la plantilla Excel que se utilizará para generar los archivos de control.
   - **Define el directorio de destino**: Escoge la carpeta donde se guardarán los archivos de control generados.
   - **Procesar archivos**: Haz clic en el botón "Procesar Archivos" para iniciar el proceso de generación de documentos.

3. **Verificación**:
   - Consulta el área de resultados en la GUI para verificar la creación exitosa de los archivos o para identificar posibles errores que necesiten atención.

## Personalización

- **Modificación de rutas y plantillas**: Puedes adaptar este script a tus necesidades específicas cambiando las rutas de los archivos o ajustando cómo se completan los campos dentro de la plantilla Excel en el código fuente. Esto te permite personalizar la entrada de datos y la presentación de los documentos de control generados.
