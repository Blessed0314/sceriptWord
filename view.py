import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from openpyxl import load_workbook

# Funciones de la interfaz

#Función para llenar el entry y guardarlo en la variable path del script
def fill_entry(path, entry):
    entry.config(state='normal')
    entry.delete(0, tk.END)
    entry.insert(0, path)
    entry.config(state='disabled')

#Función para obtener el directorio
def select_directory(entry, button):
    path = filedialog.askdirectory()
    fill_entry(path, entry)
    if entry.get():
        button.config(text="Change")

#Función para seleccionar el archivo excel
def select_excel(entry, button):
    path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
    fill_entry(path, entry)
    if entry.get():
        button.config(text="Change")
        

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
    command=lambda: select_directory(target_entry, target_button)
)
target_button.grid(row=1, column=0, padx=10, sticky='w')

# Separador
separator = ttk.Separator(ventana, orient='horizontal')
separator.grid(row=2, column=0, columnspan=2, sticky='ew', padx=5, pady=5)

# Segundo bloque, selección del path donde se guardara el documento final

#Etiqueta
label_dest = tk.Label(ventana, text="Paso 2: Selecciona el directorio donde se guardará el documento")
label_dest.grid(row=3, column=0, padx=10, columnspan=2, sticky='w')

# Field text para el mostrar el path
dest_entry = tk.Entry(ventana, width=50, state = "disabled")
dest_entry.grid(row=4, column=1, sticky='w')

# Botón para seleccionar el directorio del documento final
dest_button = tk.Button(
    ventana, 
    text="Select Path", 
    command=lambda: select_directory(dest_entry, dest_button)
)
dest_button.grid(row=4, column=0, padx=10, sticky='w')

# Separador
separator = ttk.Separator(ventana, orient='horizontal')
separator.grid(row=5, column=0, columnspan=2, sticky='ew', padx=5, pady=5)

# Tercer bloque, selección del nombre del documento

#Etiqueta
label_docName = tk.Label(ventana, text="Paso 3: Escriba el nombre del estudiante")
label_docName.grid(row=6, column=0, padx=10, columnspan=2, sticky='w')

# Field text para colocar el nombre del estudiante
labelName = tk.Label(ventana, text="Nombre:")
labelName.grid(row=7, column=0, padx=10, sticky='w')
docName_entry = tk.Entry(ventana, width=50)
docName_entry.grid(row=7, column=1, sticky='w')

# Separador
separator = ttk.Separator(ventana, orient='horizontal')
separator.grid(row=8, column=0, columnspan=2, sticky='ew', padx=5, pady=5)

# Cuarto bloque, selección del excel con los progranas académicos

#Etiqueta
label_list = tk.Label(ventana, text="Paso 4: Selecciona el excel con el listado de programas académicos")
label_list.grid(row=9, column=0, padx=10, columnspan=2, sticky='w')

# Field text para el mostrar el path
list_entry = tk.Entry(ventana, width=50, state = "disabled")
list_entry.grid(row=10, column=1, sticky='w')

# Botón para seleccionar el excel
list_button = tk.Button(
    ventana, 
    text="Select Excel", 
    command=lambda: select_excel(list_entry, list_button)
)
list_button.grid(row=10, column=0, padx=10, sticky='w')

# Separador
separator = ttk.Separator(ventana, orient='horizontal')
separator.grid(row=11, column=0, columnspan=2, sticky='ew', padx=5, pady=5)


# Ejecución de la aplicación
ventana.mainloop()



