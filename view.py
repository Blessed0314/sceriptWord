import os
import sys

import tkinter as tk
import tkinter.messagebox as messagebox
import script

from tkinter import ttk, filedialog
from tkinter import filedialog

def download_template(button):
    try:
        # Intentar descargar la plantilla
        script.create_template()
    except Exception as e:
        # Si ocurre un error, mostrar un mensaje al usuario
        tk.messagebox.showerror("Error", "No se pudo descargar la plantilla. Por favor, cierra el archivo Excel y vuelve a intentarlo.")

def select_directory(entry, button, key):
    # Solicitar al usuario que seleccione un directorio
    if key == 'excel_path':
        directory = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
    else:
        directory = filedialog.askdirectory()

    # Actualizar el campo de entrada con el directorio seleccionado
    entry.config(state="normal")
    entry.delete(0, tk.END)
    entry.insert(0, directory)
    entry.config(state="disabled")

    # Cambiar el texto del botón
    button.config(text="Change")

    # Actualizar el diccionario con el nuevo path
    paths[key] = directory

def create_input_field_button(ventana, row, label_text, button_text, command, key, entry_state="normal"):
    # Etiqueta
    label = tk.Label(ventana, text=label_text)
    label.grid(row=row, column=0, padx=10, columnspan=2, sticky='w')

    # Campo de texto
    entry = tk.Entry(ventana, width=50, state=entry_state)
    entry.grid(row=row+1, column=1, sticky='w')

    # Botón
    button = tk.Button(ventana, text=button_text)
    button.grid(row=row+1, column=0, padx=10, sticky='w')

    # Asignar la función command al botón después de que button se ha definido
    def command_with_args():
        command(entry, button, key)
    button['command'] = command_with_args

    return entry, button

def create_input_field_no_button(ventana, row, label_text, key, entry_state="normal"):
    # Etiqueta
    label = tk.Label(ventana, text=label_text)
    label.grid(row=row, column=0, padx=10, columnspan=2, sticky='w')

    # Campo de texto
    entry = tk.Entry(ventana, width=50, state=entry_state)
    entry.grid(row=row+1, column=1, sticky='w')

    # Actualizar el diccionario con el valor de la entrada cuando se modifica
    def update_dict(event=None):
        paths[key] = entry.get()
    entry.bind("<FocusOut>", update_dict)

    return entry, update_dict

def create_button(window, row, label_text, button_text, command):
    # Crear la etiqueta
    label = tk.Label(window, text=label_text)
    label.grid(row=row, column=0, padx=10, columnspan=2, sticky='w')

    # Crear el botón
    button = tk.Button(window, text=button_text)

    # Asignar la función command al botón después de que button se ha definido
    def command_with_args():
        command(button)
    button['command'] = command_with_args

    button.grid(row=row+1, column=0, columnspan=2)
    

def create_separator(ventana, row):
    # Separador
    separator = ttk.Separator(ventana, orient='horizontal')
    separator.grid(row=row, column=0, columnspan=2, sticky='ew', padx=5, pady=5)

def start_processing():
    update_docName()
    # Bloquear la ventana
    try:
        # Deshabilitar la ventana
        ventana.withdraw()

        # Llamar a script.run_script con los argumentos del diccionario
        script.run_script(paths.get('target_path'), paths.get('dest_path'), paths.get('doc_name'), paths.get('excel_path'))
    except Exception as e:
        # Mostrar un mensaje de alerta si ocurre una excepción
        messagebox.showerror("Error", "Ha ocurrido un error. Por favor, revise las recomendaciones dadas.")
    finally:
        # Habilitar la ventana de nuevo
        ventana.deiconify()

# Preparamos el logo para la ventana
if getattr(sys, 'frozen', False):
    # Estamos en un paquete congelado
    logo_path = os.path.join(sys._MEIPASS, 'logo.png')
else:
    # Estamos en un entorno normal de Python
    logo_path = os.path.join(os.path.dirname(__file__), 'logo.png')

# Crear la ventana
ventana = tk.Tk()
ventana.title("Generador de documentos")
ventana.iconphoto(False, tk.PhotoImage(file=logo_path))
ventana.geometry("395x615")
ventana.resizable(False, False)

paths = {}

# Bloque de observaciones de la aplicación
instructions = """
Recomendaciones antes de usar la aplicación:

1. La plantilla de la hoja de cálculo es para importar el historial 
    de  asignaturas con sus notas, al  descargar no modificarlo y 
    no llenar nada en la columna 'status copy'.

2. Al momento de ejecutar la aplicación, no puede estar ningún 
    documento involucrado abierto, incluída la hoja de cálculo
    y los planes programáticos.

3. Por favor, verifica que los nombres de las asignaturas en la 
    hoja de cálculo concuerden con los nombres de los archivos 
    de los planes programáticos.

Por favor, sigue las instrucciones cuidadosamente.
"""
create_separator(ventana, 1)

instructions_label = tk.Message(ventana, text=instructions, justify=tk.LEFT)
instructions_label.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

# Primer paso
template_button = create_button(ventana, 2, "Paso 1: Descargar la plantilla de excel y llenarla", "Descargar plantilla", download_template)

create_separator(ventana, 4)

# Segundo paso
excel_entry, excel_path_button = create_input_field_button(
    ventana, 5, "Paso 2: Busca y selecciona el archivo de excel ya editado", "Select Path", select_directory, "excel_path", "disabled"
)

create_separator(ventana, 7)

# Tercer paso
target_entry, target_button = create_input_field_button(
    ventana, 8, "Paso 3: Selecciona el directorio de los programas académicos", "Select Path", select_directory, "target_path", "disabled"
)

create_separator(ventana, 10)

#  Cuarto paso
dest_entry, dest_button = create_input_field_button(
    ventana, 11, "Paso 4: Selecciona el directorio donde se guardará el documento", "Select Path", select_directory, "dest_path", "disabled"
)

create_separator(ventana, 13)

# Quinto paso
docName_entry, update_docName = create_input_field_no_button(
    ventana, 14, "Paso 5: Escriba el nombre del estudiante", "doc_name", "normal"
)

create_separator(ventana, 16)

# Empezar
start_button = tk.Button(ventana, text="Empezar", command=start_processing)
start_button.grid(row=18, column=0, columnspan=2)

# Ejecución de la aplicación
ventana.mainloop()

