from flask import Blueprint, request, jsonify
from src.services.pdf_converter import parse_tokio_text, save_dataframe_to_excel
from src.utils.file_handler import save_file

upload_bp = Blueprint("upload", __name__)

@upload_bp.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado"}), 400

    file = request.files["file"]

    if file.filename == "":
        return jsonify({"error": "Nome do arquivo vazio"}), 400

    try:
        file_path = save_file(file)
        dados_header, df = parse_tokio_text(file_path)  # Certifique-se de que esta linha está correta
        excel_path = save_dataframe_to_excel(dados_header, df)

        return jsonify({
            "message": "Arquivo processado com sucesso",
            "excel_path": excel_path
        }), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500