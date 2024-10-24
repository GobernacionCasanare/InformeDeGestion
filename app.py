from flask import Flask, render_template, request, send_file, send_from_directory
import os
from openpyxl import load_workbook
from docxtpl import DocxTemplate
import json

app = Flask(__name__)
app.config['GENERATED_FOLDER'] = 'generated'  # Carpeta para los archivos generados

# Página principal para generar el informe
@app.route('/')
def index():
    return send_from_directory(os.getcwd(), 'index.html')

@app.route('/generate', methods=['POST'])
def generate_doc():
    # Cargar la plantilla .docx
    doc = DocxTemplate('1.docx')
    context = {}

    # Procesar archivos Excel que ya están en la carpeta
    excel_files = ['ejemplo1.xlsx', 'ejemplo2.xlsx']  # Añade aquí los nombres de los archivos Excel que usarás

    for filename in excel_files:
        file_path = os.path.join(os.getcwd(), filename)

        if filename.endswith('.xlsx'):
            wb = load_workbook(file_path, data_only=True)

            if 'ejemplo1' in filename:
                ws = wb.worksheets[1]  # Hoja 2
                context['ev1'] = ws['D9'].value
                context['ev2'] = ws['D10'].value

            elif 'ejemplo2' in filename:
                ws = wb.worksheets[2]  # Hoja 4
                context['ev3'] = ws['B12'].value
                ws = wb.worksheets[3]
                context['ev4'] = ws['C13'].value

    # Rellenar la plantilla con los datos
    doc.render(context)

    # Guardar el nuevo archivo .docx lleno
    output_path = os.path.join(app.config['GENERATED_FOLDER'], 'informe_lleno.docx')
    if os.path.exists(output_path):
        os.remove(output_path)
    doc.save(output_path)

    # Descargar el archivo generado
    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
