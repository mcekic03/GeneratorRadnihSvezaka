from flask import Flask, render_template, request, jsonify, send_file
import os
import pandas as pd
from generate_pdf_testing_AN import process_analiza_nastave, generate_pdf_an
from generate_pdf_testing_EDN import process_excel as process_edn, generate_pdf_edn
from generate_pdf_izvestaj_o_radu_konacno import process_excel as process_izvestaj, generate_pdf_testing_test1 as generate_pdf_izvestaj
from generate_pdf_testing_OP import process_excel as process_op, generate_pdf_testing_test1 as generate_pdf_op
import traceback
import base64
from generatorpdfkonacno import generate_pdf
from datetime import datetime
from zipfile import ZipFile

from waitress import serve

app = Flask(__name__, static_folder='static')

UPLOAD_FOLDER = 'uploads'
PDF_FOLDER = 'pdfs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PDF_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/get_available_sheets', methods=['POST'])
def get_available_sheets():
    if 'excel_file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['excel_file']
    if file.filename == '':
        return jsonify({'error': 'Fajl nije selektovan'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({'error': 'Neispravan format fajla. Učitajte Excel fajl'}), 400
    
    filename = file.filename
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    try:
        _, _, available_sheets = process_izvestaj(filepath)
        return jsonify({'sheets': available_sheets})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/generate/<report_type>', methods=['POST'])
def generate_report(report_type):
    if 'excel_file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['excel_file']
    if file.filename == '':
        return jsonify({'error': 'Fajl nije selektovan'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({'error': 'Neispravan format fajla. Učitajte Excel fajl'}), 400
    
    filename = file.filename
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    try:
        if report_type == 'an':
            # Analiza nastave
            tables_data, professor_name = process_analiza_nastave(filepath)
            if tables_data and professor_name:
                pdf_path = os.path.join(PDF_FOLDER, "analiza_nastave.pdf")
                generate_pdf_an(tables_data, pdf_path, professor_name)
                response = send_file(pdf_path, as_attachment=True, download_name="analiza_nastave.pdf", mimetype="application/pdf")
                encoded_name = base64.b64encode(professor_name.encode('utf-8')).decode('ascii')
                response.headers['Professor-Name'] = encoded_name
                return response
            else:
                return jsonify({'error': 'Greška u obradi podataka za Analizu nastave'}), 500

        elif report_type == 'edn':
            # Evidencija držanja nastave
            table_data, professor_name = process_edn(filepath)
            if table_data and professor_name:
                pdf_path = os.path.join(PDF_FOLDER, "evidencija_drzanja_nastave.pdf")
                generate_pdf_edn(table_data, pdf_path, professor_name)
                response = send_file(pdf_path, as_attachment=True, download_name="evidencija_drzanja_nastave.pdf", mimetype="application/pdf")
                encoded_name = base64.b64encode(professor_name.encode('utf-8')).decode('ascii')
                response.headers['Professor-Name'] = encoded_name
                return response
            else:
                return jsonify({'error': 'Greška u obradi podataka za Evidenciju držanja nastave'}), 500

        elif report_type == 'izvestaj':
            # Izveštaj o radu
            try:
                selected_sheet = request.form.get('selected_sheet')
                extracted_data, professor_name, _ = process_izvestaj(filepath, selected_sheet)
                if extracted_data and professor_name:
                    pdf_path = os.path.join(PDF_FOLDER, "izvestaj_o_radu.pdf")
                    generate_pdf_izvestaj(extracted_data, pdf_path, professor_name)
                    response = send_file(pdf_path, as_attachment=True, download_name="izvestaj_o_radu.pdf", mimetype="application/pdf")
                    encoded_name = base64.b64encode(professor_name.encode('utf-8')).decode('ascii')
                    response.headers['Professor-Name'] = encoded_name
                    if selected_sheet:
                        response.headers['Selected-Sheet'] = selected_sheet
                    return response
                else:
                    return jsonify({'error': 'Greška u obradi podataka za Izveštaj o radu'}), 500
            except Exception as e:
                print(f"Error processing Izvestaj data: {str(e)}")
                traceback.print_exc()
                return jsonify({'error': f'Greška u obradi podataka za Izveštaj o radu: {str(e)}'}), 500

        elif report_type == 'op':
            # Osnovni podaci
            try:
                extracted_data = process_op(filepath)
                if extracted_data:
                    # Get professor name from the first sheet's data
                    first_sheet_data = next(iter(extracted_data.values()))
                    name_data = {"first_name": "", "last_name": ""}
                    
                    # Extract professor's name from "Osnovni podaci" section
                    for row in first_sheet_data.get("Osnovni podaci", []):
                        if len(row) > 1:  # Ensure row has both label and value
                            label = str(row[0]).strip().lower()
                            value = str(row[1]).strip()
                            
                            if "ime" in label and not name_data["first_name"]:
                                name_data["first_name"] = value
                            elif "prezime" in label and not name_data["last_name"]:
                                name_data["last_name"] = value
                    
                    professor_name = f"{name_data['first_name']} {name_data['last_name']}".strip()
                    
                    pdf_path = os.path.join(PDF_FOLDER, "osnovni_podaci.pdf")
                    generate_pdf_op(extracted_data, pdf_path)
                    response = send_file(pdf_path, as_attachment=True, download_name="osnovni_podaci.pdf", mimetype="application/pdf")
                    encoded_name = base64.b64encode(professor_name.encode('utf-8')).decode('ascii')
                    response.headers['Professor-Name'] = encoded_name
                    return response
                else:
                    return jsonify({'error': 'Greška u obradi podataka za Osnovne podatke'}), 500
            except Exception as e:
                print(f"Error processing OP data: {str(e)}")
                traceback.print_exc()
                return jsonify({'error': f'Greška u obradi podataka za Osnovne podatke: {str(e)}'}), 500

        else:
            return jsonify({'error': 'Nepoznat tip izveštaja'}), 400

    except Exception as e:
        return jsonify({'error': f'Neuspelo procesiranje fajla: {str(e)}'}), 500

@app.route('/procesiranje', methods=['POST'])
def procesiranje():
    if 'excel_file' not in request.files or 'professor_name' not in request.form:
        return jsonify({'error': 'Excel fajl i ime profesora su obavezni.'}), 400
    file = request.files['excel_file']
    professor_name = request.form['professor_name']
    filename = file.filename
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)
    try:
        df = pd.read_excel(filepath)
        professor_data = df[df['Ime Predavača'] == professor_name]
        if professor_data.empty:
            return jsonify({'error': f'Nisu pronađeni podaci o profesoru sa imenom: {professor_name}'}), 404
        columns_to_send = [
            "Ime Predavača",
            "Naziv Predmeta",
            "Pozicija",
            "Tip Predavanja",
            "Nedeljni Broj Časova",
            "Broj Grupa",
            "Tip studija",
            "Odsek",
            "Ukupno casova"
        ]
        professor_data_filtered = professor_data[columns_to_send]
        pdf_name = f"{professor_name.replace(' ', '_')}_{datetime.now().strftime('Opterecenje_%Y%m%d_%H%M%S')}.pdf"
        pdf_path = os.path.join(PDF_FOLDER, pdf_name)
        generate_pdf(professor_data_filtered, pdf_path)
        return send_file(pdf_path, as_attachment=True, download_name=pdf_name, mimetype="application/pdf")
    except Exception as e:
        return jsonify({'error': f'Greška u procesiranju: {str(e)}'}), 500

@app.route('/procesiranjesvih', methods=['POST'])
def procesiranjesvih():
    if 'excel_file_svih' not in request.files:
        return jsonify({'error': 'Excel fajl je obavezan.'}), 400
    file = request.files['excel_file_svih']
    filename = file.filename
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)
    try:
        df = pd.read_excel(filepath)
        if 'Ime Predavača' not in df.columns:
            return jsonify({'error': "Excel fajl mora da sadrži 'Ime Predavača' kolonu."}), 400
        unique_professors = df['Ime Predavača'].dropna().unique()
        pdf_files = []
        for professor_name in unique_professors:
            professor_data = df[df['Ime Predavača'] == professor_name]
            if not professor_data.empty:
                columns_to_send = [
                    "Ime Predavača",
                    "Naziv Predmeta",
                    "Pozicija",
                    "Tip Predavanja",
                    "Nedeljni Broj Časova",
                    "Broj Grupa",
                    "Tip studija",
                    "Odsek",
                    "Ukupno casova"
                ]
                professor_data_filtered = professor_data[columns_to_send]
                pdf_name = f"{professor_name.replace(' ', '_')}_{datetime.now().strftime('Opterecenje_%Y%m%d_%H%M%S')}.pdf"
                pdf_path = os.path.join(PDF_FOLDER, pdf_name)
                generate_pdf(professor_data_filtered, pdf_path)
                pdf_files.append(pdf_name)
        if not pdf_files:
            return jsonify({'error': 'Nije generisan nijedan PDF. Proverite učitane podatke.'}), 400
        zip_filename = f"professors_pdfs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        zip_filepath = os.path.join(PDF_FOLDER, zip_filename)
        with ZipFile(zip_filepath, 'w') as zipf:
            for pdf_file in pdf_files:
                zipf.write(os.path.join(PDF_FOLDER, pdf_file), pdf_file)
        return send_file(zip_filepath, as_attachment=True, download_name=zip_filename, mimetype="application/zip")
    except Exception as e:
        return jsonify({'error': f'Greška u procesiranju svih: {str(e)}'}), 500

if __name__ == "__main__": 
    #app.run(debug=True, port=1000)
    serve(app, host='0.0.0.0', port=80)