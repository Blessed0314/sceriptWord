from docx import Document
from pathlib import Path
'''
Esta primera parte del código es para generar un documento de word 
la idea es que el nombre del documento lo tome de la interfaz grafica, asi mismo
como también el directorio donde se va a guardar. 
'''

finalDoc = Document() # Creamos el objeto documento doonse se almacenara toda la info
docName = "blanco" # Nombre del documento 
docDir = "C:/Users/user/Desktop/Trabajos/scriptWord/DocumentosGenerados/"+ docName + ".docx"


'''
Directorio donde se obtendrán los documentos para ser fusionados, la
idea es que esta carpeta se pueda seleccionar desde la interfaz
'''
targetDocs = "C:/Users/user/Desktop/Trabajos/scriptWord/DocumentosObjetivos"


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
