import os

UPLOAD_FOLDER = "uploads"

def save_file(file):
    """
    Salva o arquivo enviado no servidor.
    :param file: Arquivo enviado pelo cliente.
    :return: Caminho do arquivo salvo.
    """
    try:
        if not os.path.exists(UPLOAD_FOLDER):
            os.makedirs(UPLOAD_FOLDER)

        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)
        return file_path
    except Exception as e:
        raise Exception(f"Erro ao salvar arquivo: {e}")