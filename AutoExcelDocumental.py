import pandas as pd
from openpyxl import load_workbook
import os
import re
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import PIL.Image
import PIL.ImageTk

def procesar_archivos():
    ruta_archivo_fuente = entry_fuente.get()
    ruta_archivo_control_plantilla = entry_plantilla.get()
    ruta_destino_controles = entry_destino.get()

    # Leer el archivo Excel fuente
    df_fuente = pd.read_excel(ruta_archivo_fuente, header=None)

    # Limpiar el área de texto antes de empezar
    resultado_text.delete(1.0, tk.END)

    # Iterar sobre cada fila del DataFrame
    for index, row in df_fuente.iterrows():
        # Verificar si hay algún texto en la fila
        if any(row.notna()):
            # Obtener la información de la fila
            NombreArchivo = row[0]
            FechaInicio = row[1]
            FechaFinal = row[2]
            nombre_carpeta = row[0]

            # Formatear el nombre de la carpeta eliminando caracteres no permitidos
            nombre_carpeta = re.sub(r'[\\/*?:"<>|]', '', str(nombre_carpeta))

            # Construir la ruta del nuevo archivo "control"
            ruta_nuevo_control = os.path.join(ruta_destino_controles, f'0 {nombre_carpeta}.xlsx')

            # Verificar si el archivo ya existe en la ruta especificada
            if os.path.exists(ruta_nuevo_control):
                resultado_text.insert(tk.END, f"{index + 1}). El archivo control en la ruta {ruta_nuevo_control} ya existe.\n\n")
            else:
                # Cargar la plantilla "control" utilizando openpyxl
                wb_control = load_workbook(ruta_archivo_control_plantilla)
                ws_control = wb_control.active

                # Modificar las celdas con las fechas
                ws_control['C10'] = NombreArchivo
                ws_control['A13'] = FechaInicio
                ws_control['A14'] = FechaInicio
                ws_control['A15'] = FechaFinal

                # Guardar el nuevo archivo "control" en la ruta de destino
                wb_control.save(ruta_nuevo_control)

                resultado_text.insert(tk.END, f"{index + 1}). Se ha creado el nuevo archivo control en la ruta {ruta_nuevo_control}\n")

    resultado_text.insert(tk.END, "Proceso completado. Los archivos se guardaron en las rutas especificadas.\n")

def despedida():
    resultado_text.delete(1.0, tk.END)
    resultado_text.insert(tk.END, "Hasta pronto, Feliz día  :)\n")
    root.after(1000, root.destroy)

# Crear la interfaz gráfica
root = tk.Tk()
root.title("AUTO EXCEL DOCUMENTAL CEAI")

#Configurar el logo 
ruta_logo = 'C:\\Users\\usuario\\Desktop\\py\\Logo\\logoSena.png'
logo_pillow = PIL.Image.open(ruta_logo)
logo_pillow = logo_pillow.resize((100, 100))  # Corregir aquí
logo_image = PIL.ImageTk.PhotoImage(logo_pillow)

logo_label = tk.Label(root, text="AUTO EXCEL DOCUMENTAL CEAI", image=logo_image, compound="top")
logo_label.grid(row=0, column=0, columnspan=2, pady=(8, 0))

# Etiquetas y campos de entrada
tk.Label(root, text="Ruta del archivo fuente:").grid(row=1, column=0, sticky=tk.E)
tk.Label(root, text="Ruta de la plantilla control:").grid(row=2, column=0, sticky=tk.E)
tk.Label(root, text="Ruta de destino para controles:").grid(row=3, column=0, sticky=tk.E)

entry_fuente = ttk.Entry(root, width=50)
entry_plantilla = ttk.Entry(root, width=50)
entry_destino = ttk.Entry(root, width=50)

entry_fuente.grid(row=1, column=1, padx=5, pady=5) 
entry_plantilla.grid(row=2, column=1, padx=5, pady=5)
entry_destino.grid(row=3, column=1, padx=5, pady=5)

# Botón para procesar archivos
procesar_button = ttk.Button(root, text="Procesar Archivos", command=procesar_archivos, style="Green.TButton")
procesar_button.grid(row=4, column=0, columnspan=2, pady=(10, 5))

# Área de texto para mostrar el resultado con barra de desplazamiento
resultado_text = tk.Text(root, height=10, width=60, wrap=tk.WORD)
resultado_text.grid(row=5, column=0, columnspan=2, padx=5, pady=5)

scrollbar = tk.Scrollbar(root, command=resultado_text.yview)
scrollbar.grid(row=5, column=2, sticky='nsew')
resultado_text['yscrollcommand'] = scrollbar.set

# Botón de salida
salir_button = ttk.Button(root, text="Salir", command=despedida, style="Red.TButton")
salir_button.grid(row=6, column=0, columnspan=2, pady=5)

# Configurar el estilo para los botones
style = ttk.Style()
style.configure("Green.TButton", foreground="black", background="#1FF436")
style.configure("Red.TButton", foreground="black", background="#FF0000")

root.mainloop()
