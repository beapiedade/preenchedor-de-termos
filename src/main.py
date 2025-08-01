from extractor import Extractor
from factory import Factory
import os
import shutil
import uuid
import zipfile
from flask import Flask, render_template, request, send_file, after_this_request

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
GENERATED_FOLDER = "generated"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/processar', methods=['POST'])
def processar_arquivos():
    if "excel_file" not in request.files or "template_file" not in request.files:
        return "Erro: Arquivos não enviados corretamente.", 400

    excel_file = request.files["excel_file"]
    template_file = request.files["template_file"]

    if excel_file.filename == "" or template_file.filename == "":
        return "Erro: Nenhum arquivo selecionado.", 400

    session_id = str(uuid.uuid4())
    session_upload_path = os.path.join(UPLOAD_FOLDER, session_id)
    session_generated_path = os.path.join(GENERATED_FOLDER, session_id)
    os.makedirs(session_upload_path, exist_ok=True)
    os.makedirs(session_generated_path, exist_ok=True)
    
    excel_path = os.path.join(session_upload_path, excel_file.filename)
    template_path = os.path.join(session_upload_path, template_file.filename)
    excel_file.save(excel_path)
    template_file.save(template_path)

    try:
        data = Extractor.from_path(excel_path)
        final_data, data_errors = Extractor.check_data(data)

        if not final_data:
            return "Nenhum dado válido encontrado na planilha para gerar documentos.", 400

        for item in final_data:
            Factory.create_document(session_generated_path, template_path, item)

        zip_path = os.path.join(GENERATED_FOLDER, f"{session_id}.zip")
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for root, _, files in os.walk(session_generated_path):
                for file in files:
                    zipf.write(os.path.join(root, file), arcname=file)

        @after_this_request
        def cleanup(response):
            try:
                shutil.rmtree(session_upload_path)
                shutil.rmtree(session_generated_path)
                os.remove(zip_path)
            except Exception as e:
                print(f"Erro ao limpar arquivos: {e}")
            return response

        return send_file(zip_path, as_attachment=True, download_name="documentos_gerados.zip")

    except Exception as e:
        return f"Ocorreu um erro inesperado: {str(e)}", 500

if __name__ == "__main__":
    app.run(debug=True)