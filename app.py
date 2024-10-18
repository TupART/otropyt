from flask import Flask, request, send_file, render_template
import pandas as pd
import tempfile
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    uploaded_file = request.files['file']
    if uploaded_file.filename != '':
        # Leer el archivo .xlsx
        df = pd.read_excel(uploaded_file, header=1)

        # Validación: Si hay filas con menos de 19 columnas
        if df.shape[1] < 19:
            return "Error: El archivo debe tener al menos 19 columnas.", 400

        # Preparar la plantilla
        template_path = 'PlantillaSTEP4.xlsx'
        template = pd.read_excel(template_path, header=6)

        # Rellenar datos en la plantilla
        for index, row in df.iterrows():
            # Rellenar Name
            template.at[index + 7, 'Name'] = row['Name']
            # Rellenar Surname
            template.at[index + 7, 'Surname'] = row['Surname']
            # Rellenar Primary email
            template.at[index + 7, 'Primary email'] = row['E-mail']
            
            # Rellenar Primary phone
            if row['Va a ser PCC?'] == 'Y':
                if row['Market'] == 'DACH':
                    template.at[index + 7, 'Primary phone'] = '/+4940210918145 /+43122709858 /+41445295828'
                elif row['Market'] == 'France':
                    template.at[index + 7, 'Primary phone'] = '/+33180037979'
                elif row['Market'] == 'Spain':
                    template.at[index + 7, 'Primary phone'] = '/+34932952130'
                elif row['Market'] == 'Italy':
                    template.at[index + 7, 'Primary phone'] = '/+390109997099'
            else:
                template.at[index + 7, 'Primary phone'] = ''
            
            # Rellenar Workgroup
            if row['Va a ser PCC?'] == 'Y':
                if row['Market'] == 'DACH':
                    template.at[index + 7, 'Workgroup'] = 'D_PCC'
                elif row['Market'] == 'France':
                    template.at[index + 7, 'Workgroup'] = 'F_PCC'
                elif row['Market'] == 'Spain':
                    template.at[index + 7, 'Workgroup'] = 'E_PCC'
                elif row['Market'] == 'Italy':
                    template.at[index + 7, 'Workgroup'] = 'I_PCC'
            elif row['Va a ser PCC?'] == 'N':
                if row['Market'] == 'DACH':
                    template.at[index + 7, 'Workgroup'] = 'D_Outbound'
                elif row['Market'] == 'France':
                    template.at[index + 7, 'Workgroup'] = 'F_Outbound'
                elif row['Market'] == 'Spain':
                    template.at[index + 7, 'Workgroup'] = 'E_Outbound'
            elif row['Va a ser PCC?'] == 'TL':
                if row['Market'] == 'DACH':
                    template.at[index + 7, 'Workgroup'] = 'D_PCC'
                elif row['Market'] == 'France':
                    template.at[index + 7, 'Workgroup'] = 'F_PCC'
                elif row['Market'] == 'Spain':
                    template.at[index + 7, 'Workgroup'] = 'E_PCC'
                elif row['Market'] == 'Italy':
                    template.at[index + 7, 'Workgroup'] = 'I_PCC'
            elif row['Va a ser PCC?'] == 'DS':
                if row['Market'] == 'DACH':
                    template.at[index + 7, 'Workgroup'] = 'D_Outbound'
                elif row['Market'] == 'France':
                    template.at[index + 7, 'Workgroup'] = 'F_Outbound'
                elif row['Market'] == 'Spain':
                    template.at[index + 7, 'Workgroup'] = 'E_Outbound'

            # Rellenar Team
            if row['Va a ser PCC?'] == 'Y':
                if row['Market'] == 'DACH':
                    template.at[index + 7, 'Team'] = 'Team_D_CCH_PCC_1'
                elif row['Market'] == 'France':
                    template.at[index + 7, 'Team'] = 'Team_F_CCH_PCC_1'
                elif row['Market'] == 'Spain':
                    template.at[index + 7, 'Team'] = 'Team_E_CCH_PCC_1'
                elif row['Market'] == 'Italy':
                    template.at[index + 7, 'Team'] = 'Team_I_CCH_PCC_1'
            elif row['Va a ser PCC?'] == 'N':
                if row['Market'] == 'DACH':
                    template.at[index + 7, 'Team'] = 'Team_D_CCH_B2C_1'
                elif row['Market'] == 'France':
                    template.at[index + 7, 'Team'] = 'Team_F_CCH_B2C_1'
                elif row['Market'] == 'Spain':
                    template.at[index + 7, 'Team'] = 'Team_E_CCH_B2C_1'
            elif row['Va a ser PCC?'] == 'TL':
                if row['Market'] == 'DACH':
                    template.at[index + 7, 'Team'] = 'Team_D_CCH_PCC_1'
                elif row['Market'] == 'France':
                    template.at[index + 7, 'Team'] = 'Team_F_CCH_PCC_1'
                elif row['Market'] == 'Spain':
                    template.at[index + 7, 'Team'] = 'Team_E_CCH_PCC_1'
                elif row['Market'] == 'Italy':
                    template.at[index + 7, 'Team'] = 'Team_I_CCH_PCC_1'
            elif row['Va a ser PCC?'] == 'DS':
                if row['Market'] == 'DACH':
                    template.at[index + 7, 'Team'] = 'Team_D_CCH_B2C_1'
                elif row['Market'] == 'France':
                    template.at[index + 7, 'Team'] = 'Team_F_CCH_B2C_1'
                elif row['Market'] == 'Spain':
                    template.at[index + 7, 'Team'] = 'Team_E_CCH_B2C_1'

            # Rellenar Is PCC
            if row['Va a ser PCC?'] == 'Y':
                template.at[index + 7, 'Is PCC'] = 'Y'
            elif row['Va a ser PCC?'] == 'N':
                template.at[index + 7, 'Is PCC'] = 'N'
            elif row['Va a ser PCC?'] == 'TL':
                template.at[index + 7, 'Is PCC'] = 'N'
            elif row['Va a ser PCC?'] == 'DS':
                template.at[index + 7, 'Is PCC'] = 'N'

            # Rellenar CTI User
            template.at[index + 7, 'CTI User'] = row['B2E User Name']
            # Rellenar TTG UserID 1
            template.at[index + 7, 'TTG UserID 1'] = row['B2E User Name']

            # Rellenar Campaign Level
            if row['Va a ser PCC?'] == 'Y':
                if row['Market'] in ['DACH', 'France', 'Spain', 'Italy']:
                    template.at[index + 7, 'Campaign Level'] = 'Agent'
            elif row['Va a ser PCC?'] == 'N':
                if row['Market'] in ['DACH', 'France', 'Spain']:
                    template.at[index + 7, 'Campaign Level'] = 'Agent'
            elif row['Va a ser PCC?'] == 'TL':
                if row['Market'] in ['DACH', 'France', 'Spain', 'Italy']:
                    template.at[index + 7, 'Campaign Level'] = 'Team Leader'
            elif row['Va a ser PCC?'] == 'DS':
                if row['Market'] in ['DACH', 'France', 'Spain']:
                    template.at[index + 7, 'Campaign Level'] = 'Agent'

        # Guardar el archivo resultante en un archivo temporal
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        template.to_excel(temp_file.name, index=False)
        
        return send_file(temp_file.name, as_attachment=True, download_name='resultado.xlsx')

if __name__ == '__main__':
    app.run(debug=True)
