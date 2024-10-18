from flask import Flask, render_template, request, send_file
import pandas as pd
import openpyxl
from werkzeug.utils import secure_filename
import os
import tempfile

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Subir archivo
        if 'file' not in request.files:
            return 'No file part'
        file = request.files['file']
        if file.filename == '':
            return 'No selected file'
        if file:
            filename = secure_filename(file.filename)
            
            # Guardar el archivo en un archivo temporal
            with tempfile.NamedTemporaryFile(delete=False) as temp_file:
                file_path = temp_file.name
                file.save(file_path)
            
            # Leer el archivo con pandas
            df = pd.read_excel(file_path, header=1)  # Nombres de columnas en la fila 2 (índice 1)
            
            # Seleccionar todas las filas para mostrar en la tabla
            data = df[['Name', 'Surname', 'E-mail', 'Market', 'Va a ser PCC?', 'B2E User Name']].to_dict(orient='records')
            
            return render_template('index.html', data=data)
    
    return render_template('index.html', data=None)

@app.route('/process', methods=['POST'])
def process():
    selected_rows = request.form.getlist('rows')  # Obtener filas seleccionadas

    # Cargar el archivo original
    file_path = os.path.join(tempfile.gettempdir(), os.listdir(tempfile.gettempdir())[0])
    df = pd.read_excel(file_path, header=1)

    # Cargar la plantilla 'PlantillaSTEP4.xlsx'
    plantilla = 'PlantillaSTEP4.xlsx'
    wb = openpyxl.load_workbook(plantilla)
    ws = wb.active

    # Procesar las filas seleccionadas y rellenar la plantilla
    for row in selected_rows:
        idx = int(row)  # Convertir el índice de string a entero

        name = df.iloc[idx]['Name']
        surname = df.iloc[idx]['Surname']
        email = df.iloc[idx]['E-mail']
        market = df.iloc[idx]['Market']
        pcc_status = df.iloc[idx]['Va a ser PCC?']
        b2e_username = df.iloc[idx]['B2E User Name']

        row_num = 7 + idx  # Comenzar desde la fila 7
        ws[f'C{row_num}'] = name
        ws[f'D{row_num}'] = surname
        ws[f'E{row_num}'] = email

        # Condiciones basadas en "Market" y "PCC Status"
        if pcc_status == 'Y' and market == 'DACH':
            ws[f'F{row_num}'] = "/+4940210918145"
            ws[f'G{row_num}'] = "D_PCC"
            ws[f'H{row_num}'] = "Team_D_CCH_PCC_1"
        # Agregar más condiciones para otros mercados y estados PCC...

        ws[f'L{row_num}'] = "Y" if pcc_status == 'Y' else "N"
        ws[f'Q{row_num}'] = b2e_username
        ws[f'R{row_num}'] = b2e_username
        ws[f'V{row_num}'] = "Agent" if pcc_status in ['Y', 'N', 'DS'] else "Team Leader"

    # Guardar el archivo actualizado
    output_file = os.path.join(tempfile.gettempdir(), 'PlantillaSTEP4_Rellenada.xlsx')
    wb.save(output_file)

    # Enviar el archivo descargable
    return send_file(output_file, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
