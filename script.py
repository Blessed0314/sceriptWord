import openpyxl
import os
import subprocess
from docx import Document
from pathlib import Path
'''
Función para obtener almacenar en variables los directorios de las carpetas y para crear el documento donde se guardara el archivo final.
'''
def get_folders(targetDir, finalDir, docName):

  global finalDoc, docDir, targetDocs #Creamos las variables que se usaran en el script en forma global
  finalDoc = Document() # Creamos el objeto documento donde se almacenara toda la info
  docDir = targetDir + docName + ".docx" # Directorio donde se guardará el documento final
  targetDocs = targetDir #Directorio donde están los programas academicos

#Funcion que crea un excel base para que se copien las asignaturas y las notas de las mismas
def createTemplate ():
  template = openpyxl.Workbook()#Creamos el objeto o documento de excel
  sheet1 = template.active #Cargamos la primera hoja del excel
  sheet1['A1'] = 'Asignatura' #Colocamos el titulo a la primera celda
  sheet1['A2'] = 'Nota' #Colocamos el titulo a la segunda celda
  excelName = 'Plantilla.xlsx'
  template.save(excelName)
  template.close()
  
  # Abrir el archivo Excel recién creado
  # Comprobar si estamos en Windows
  if os.name == 'nt':  
    subprocess.Popen(['start', 'excel', excelName], shell=True)
  # Comprobar si estamos en sistemas basados en Unix/Linux
  elif os.name == 'posix':  
    subprocess.Popen(['xdg-open', excelName])

# Función para copiar párrafos
def copy_paragraphs(source_paragraphs, target_document):
    for source_paragraph in source_paragraphs:
        new_paragraph = target_document.add_paragraph()
        for run in source_paragraph.runs:
            new_run = new_paragraph.add_run(run.text)
            new_run.font.size = run.font.size  # Mantener el tamaño de la letra
            new_run.bold = run.bold
            new_run.italic = run.italic
            new_run.underline = run.underline
            

# Función para copiar tablas
def copy_tables(source_tables, target_document):
    for source_table in source_tables:
        new_table = target_document.add_table(rows=len(source_table.rows), cols=len(source_table.columns))
        for i, row in enumerate(source_table.rows):
            for j, cell in enumerate(row.cells):
                new_table.cell(i, j).text = cell.text
                new_table.cell(i, j).paragraphs[0].runs[0].font.size = cell.paragraphs[0].runs[0].font.size

# Consolidación de documentos de Word
root = Path(targetDocs)

for doc_path in root.rglob('[!.]*.docx'):
    doc = Document(doc_path)
    copy_paragraphs(doc.paragraphs, finalDoc)
    copy_tables(doc.tables, finalDoc)
    finalDoc.add_page_break() # Inserta un salto de pagina después de cada documento
    
# Guardar el documento final
finalDoc.save(docDir)


  