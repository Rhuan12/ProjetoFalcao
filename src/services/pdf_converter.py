import PyPDF2
import pandas as pd
import re
import pytesseract
from pdf2image import convert_from_path
from PIL import Image


def convert_pdf_to_text(pdf_path):
    """
    Converte um PDF em texto.
    1. Tenta extrair com PyPDF2.
    2. Se não encontrar nada, usa OCR (pytesseract).
    """
    text = ""

    # --- Tentativa 1: PyPDF2 ---
    try:
        with open(pdf_path, "rb") as pdf_file:
            reader = PyPDF2.PdfReader(pdf_file)
            for page in reader.pages:
                if page.extract_text():
                    text += page.extract_text() + "\n"
    except Exception as e:
        print(f"[ERRO] PyPDF2 falhou: {e}")

    # --- Tentativa 2: OCR ---
    if not text.strip():
        print("⚠️ PDF sem texto. Usando OCR...")
        images = convert_from_path(pdf_path, dpi=300)
        for img in images:
            text += pytesseract.image_to_string(img, lang="por") + "\n"

    return text


def extract_field(patterns, text):
    """
    Procura por uma lista de padrões regex e retorna o valor encontrado.
    """
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()  # Retorna o texto após os dois pontos
    return ""


def parse_tokio_text(text: str):
    """
    Extrai dados de um PDF da Tokio Marine (exemplo dado).
    Retorna (dados_header, df_veiculos).
    """

    # -------- Cabeçalho --------
    dados_header = {
        "NOME DO CLIENTE": extract_field([r"Proprietário[:\s]*(.*)"], text),
        "CNPJ": extract_field([r"CNPJ[:\s]*(.*)"], text),
        "APÓLICE": extract_field([r"Nr Apólice.*?(\d+)"], text),
        "VIGÊNCIA": extract_field([r"Venc Apólice.*?:\s*([\d/]+)"], text),
    }

    # -------- Dados Veículo --------
    dados_veiculo = {
        "DESCRIÇÃO DO ITEM": extract_field([r"Descrição do Item - (.*)"], text),
        "CEP DE PERNOITE DO VEÍCULO": extract_field([r"CEP de Pernoite do Veículo:\s*(.*)"], text),
        "TIPO DE UTILIZAÇÃO": extract_field([r"Tipo de utilização:\s*(.*)"], text),
        "VEÍCULO": extract_field([r"Veículo:\s*(.*)"], text),
        "ANO MODELO": extract_field([r"Ano Modelo:\s*(\d{4})"], text),
        "CHASSI": extract_field([r"Chassi:\s*(.*)"], text),
        "PLACA": extract_field([r"Placa:\s*(.*)"], text),
        "COMBUSTÍVEL": extract_field([r"Combustível:\s*(.*)"], text),
        "LOTAÇÃO VEÍCULO": extract_field([r"Lotação Veículo:\s*(.*)"], text),
        "VEÍCULO 0KM": extract_field([r"Veículo 0km:\s*(.*)"], text),
        "VEÍCULO BLINDADO": extract_field([r"Veículo Blindado:\s*(.*)"], text),
        "VEÍCULO COM KIT GÁS": extract_field([r"Veículo com Kit Gás:\s*(.*)"], text),
        "TIPO DE CARROCERIA": extract_field([r"Tipo de Carroceria:\s*(.*)"], text),
        "ISENÇÃO FISCAL": extract_field([r"Isenção Fiscal:\s*(.*)"], text),
        "PROPRIETÁRIO": extract_field([r"Proprietário:\s*(.*)"], text),
        "FIPE": extract_field([r"Fipe:\s*(.*)"], text),
        "TIPO DE SEGURO": "Renovação Tokio sem sinistro",  # Valor fixo
        "NR APÓLICE CONGENERE": extract_field([r"Nr Apólice Congenere:\s*(.*)"], text),
        "NOME DA CONGENERE": extract_field([r"Nome da Congenere:\s*(.*)"], text),
        "VENC APÓLICE CONGENERE": extract_field([r"Venc Apólice Cong.: (.*)"], text),
        "CLASSE DE BÔNUS": extract_field([r"Classe de Bônus:\s*(.*)"], text),
        "CÓDIGO DE IDENTIFICAÇÃO (CI)": extract_field([r"Código de Identificação \(CI\):\s*(.*)"], text),
        "KM DE REBOQUE": extract_field([r"Km de Reboque:\s*(.*)"], text),
        "KM (ADICIONAL)": extract_field([r"km\(Adicional\):\s*(.*)"], text),
    }

    colunas = list(dados_veiculo.keys())
    df = pd.DataFrame([dados_veiculo], columns=colunas)

    print("Dados extraídos:")
    print(dados_veiculo)

    return dados_header, df
    """
    Extrai dados de um PDF da Tokio Marine (exemplo dado).
    Retorna (dados_header, df_veiculos).
    """

def save_dataframe_to_excel(dados_header, df, output_path="output.xlsx"):
    """
    Salva cabeçalho e tabela de veículos em planilha Excel.
    """
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame([dados_header]).to_excel(writer, sheet_name="Dados Gerais", index=False)
        df.to_excel(writer, sheet_name="Veículos", index=False)
    return output_path
