import tkinter as tk
from tkinter import filedialog
from tkinter import ttk

# Funciones de la interfaz

#Función para llenar el entry y guardarlo en la variable path del script
def fill_entry(path, entry):
    entry.config(state='normal')
    entry.delete(0, tk.END)
    entry.insert(0, path)
    entry.config(state='disabled')

#Función para obtener el directorio
def select_directory(entry):
    path = filedialog.askdirectory()
    fill_entry(path, entry)
    if entry.get():
        change_button.grid(row=1, column=0, padx=10, sticky='w')
        target_button.grid_forget()

# Crear una ventana 
ventana = tk.Tk()
icono = tk.PhotoImage(file="logo.png")

# Titulo e icono de la ventana
ventana.title("Generador de Documentos")
ventana.iconphoto(True, icono)

# Parametros de la ventana
ventana.geometry("405x500")
ventana.columnconfigure(1, weight=1)
ventana.resizable(False, False)

# Primer bloque, selección del path de los archivos a fusionar

# Etiqueta
label_target = tk.Label(ventana, text="Paso 1: Selecciona el directorio de los programas académicos")
label_target.grid(row=0, column=0, padx=10, columnspan=2, sticky='w')

# Field text para el mostrar el path
target_entry = tk.Entry(ventana, width=50, state = "disabled")
target_entry.grid(row=1, column=1, sticky='w')

# Botón para seleccionar el directorio de los archivos a fusionar
target_button = tk.Button(
    ventana, 
    text="Select Path", 
    command=lambda: select_directory(target_entry)
)
target_button.grid(row=1, column=0, padx=10, sticky='w')

# Botón para cambiar el directorio seleccionado
change_button = tk.Button(
    ventana, 
    text="Change Path", 
    command=lambda: select_directory(target_entry)
)
# Separador
separator = ttk.Separator(ventana, orient='horizontal')
separator.grid(row=2, column=0, columnspan=2, sticky='ew', padx=5, pady=5)

# Segundo bloque, selección del path donde se guardara el documento final


# Ejecución de la aplicación
ventana.mainloop()



