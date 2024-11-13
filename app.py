import os
import re
import PyPDF2
import pandas as pd
from flask import Flask, request, send_file, jsonify, render_template
from werkzeug.utils import secure_filename
import shutil  # Importando shutil para apagar a pasta

app = Flask(__name__)

# Diretório absoluto para armazenar os uploads
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)  # Cria o diretório uploads, se não existir

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['ALLOWED_EXTENSIONS'] = {'pdf'}

# Função para verificar se o arquivo tem a extensão permitida
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

# Função para gerar o Excel
def generate_excel(file1):
    folder = app.config['UPLOAD_FOLDER']  # Usar o diretório absoluto de uploads
    os.chdir(folder)  # Mudar para o diretório de uploads

    # Função interna para dividir o PDF em páginas
    def split_pdf(file_path, output_dir):
        with open(file_path, 'rb') as f:
            pdf_reader = PyPDF2.PdfReader(f)
            file_list = []
            for page_num in range(len(pdf_reader.pages)):
                pdf_writer = PyPDF2.PdfWriter()
                pdf_writer.add_page(pdf_reader.pages[page_num])
                output_filename = os.path.join(
                    output_dir, f'page_{page_num + 1}.pdf')
                with open(output_filename, 'wb') as out:
                    pdf_writer.write(out)
                file_list.append(output_filename)
        return file_list

    # Função para extrair texto do PDF
    def extract_text_from_pdf(pdf_path):
        with open(pdf_path, 'rb') as f:
            pdf_reader = PyPDF2.PdfReader(f)
            text = ""
            for page_num in range(len(pdf_reader.pages)):
                text += pdf_reader.pages[page_num].extract_text()
        return text

    # Função para limpar o nome
    def clean_name(name_parts):
        cleaned_name = []
        for part in name_parts:
            if not any(char.isdigit() for char in part) and part not in ["-", "CEFAAN", "CODEMA", "SVCVPM"]:
                cleaned_name.append(part)
        return " ".join(cleaned_name)

    flist = split_pdf(file1, folder)

    # Gera a tabela com os dados de interesse
    file_paths = [os.path.join(folder, f) for f in os.listdir(folder) if f.endswith('.pdf')]

    column_names = ["Nome", "NIP", "AUX TRANSP", "AUX TRAN AC", "DES AUX TRAN", "Soldo", "END"]

    def extract_soldo(lines):
        for line in lines:
            match = re.search(r'SOLDO.*?([\d.]+,\d{2})', line)
            if match:
                return match.group(1)
        return "0"

    def extract_aux_transp(lines):
        for line in lines:
            match = re.search(r'AUX TRANSP.*?([\d.]+,\d{2})', line)
            if match:
                return match.group(1)
        return "0"

    def extract_aux_tran_ac(lines):
        for line in lines:
            match = re.search(r'AUX TRAN AC.*?([\d.]+,\d{2})', line)
            if match:
                return match.group(1)
        return "0"

    def extract_des_aux_tran(lines):
        for line in lines:
            match = re.search(r'DES AUX TRAN.*?([\d.]+,\d{2})', line)
            if match:
                return match.group(1)
        return "0"

    Tab = pd.DataFrame(columns=column_names)

    for file_path in file_paths:
        text = extract_text_from_pdf(file_path)
        lines = text.split('\n')
        lines = [line.strip() for line in lines]

        new_row = {}

        # Pega o Nome
        for i, line in enumerate(lines):
            if "PAGADORIA DE PESSOAL DA MARINHA" in line:
                id_line = lines[i + 2]
                id_parts = id_line.split()
                if len(id_parts) > 1:
                    NOME = clean_name(id_parts[1:])
                    new_row["Nome"] = NOME
                break

        # Pega o NIP
        for i, line in enumerate(lines):
            if "PAGADORIA DE PESSOAL DA MARINHA" in line:
                id_line = lines[i + 3]
                id_parts = id_line.split()
                if len(id_parts) > 1:
                    full_number = id_parts[len(id_parts)-3]
                    CPF = full_number[:11]
                    NIP = full_number[11:]
                    new_row["NIP"] = NIP
                break

        # Pega o ENDER
        for i, line in enumerate(lines):
            if "PAGADORIA DE PESSOAL DA MARINHA" in line:
                id_line = lines[i + 2]
                id_parts = id_line.split()
                if len(id_parts) > 1:
                    ENDER = id_parts[0]
                    new_row["END"] = ENDER
                break

        # Pega auxilio transporte
        AUX_TRANSP = extract_aux_transp(lines)
        new_row["AUX TRANSP"] = AUX_TRANSP

        # Pega o valor do auxílio transporte AC
        AUX_TRAN_AC = extract_aux_tran_ac(lines)
        new_row["AUX TRAN AC"] = AUX_TRAN_AC

        # Pega o valor do desconto do aux transp
        DES_AUX_TRAN = extract_des_aux_tran(lines)
        new_row["DES AUX TRAN"] = DES_AUX_TRAN

        # Pega o valor do Soldo
        SOLDO = extract_soldo(lines)
        new_row["Soldo"] = SOLDO

        if new_row:
            Tab = pd.concat([Tab, pd.DataFrame([new_row])], ignore_index=True)

    if not Tab.empty:
        output_file = os.path.join(folder, "Transporte.xlsx")
        Tab.to_excel(output_file, index=False)

        # Retorna o arquivo para o usuário antes de excluir
        return send_file(output_file, as_attachment=True)

    else:
        return jsonify({"error": "No data extracted from PDF."}), 400

# Rota para a página inicial, com o formulário de upload
@app.route('/')
def index():
    return render_template('index.html')

# Rota para upload de arquivo PDF e gerar o Excel
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        output_file = generate_excel(filepath)
        if output_file:
            # Limpar a pasta de uploads após o envio do arquivo
            for file_path in os.listdir(app.config['UPLOAD_FOLDER']):
                file_path_full = os.path.join(app.config['UPLOAD_FOLDER'], file_path)
                try:
                    if os.path.isdir(file_path_full):
                        shutil.rmtree(file_path_full)  # Remove diretórios
                    else:
                        os.remove(file_path_full)  # Remove arquivos
                except Exception as e:
                    print(f"Erro ao remover arquivo: {e}")
            return output_file
        else:
            return jsonify({"error": "No data extracted from PDF."}), 400
    else:
        return jsonify({"error": "Invalid file type."}), 400

if __name__ == '__main__':
    app.run(debug=True)
