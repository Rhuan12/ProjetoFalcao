import streamlit as st
import PyPDF2
import pandas as pd
import re
import os
from io import BytesIO
import tempfile
import pytesseract
from pdf2image import convert_from_path
from PIL import Image

# Configuração da página
st.set_page_config(
    page_title="Conversor de Apólices Tokio Marine",
    page_icon="📄",
    layout="wide"
)

def extract_text_from_pdf(pdf_file):
    """
    Extrai texto de um arquivo PDF usando PyPDF2 e OCR como fallback
    """
    try:
        # Cria um arquivo temporário
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(pdf_file.read())
            tmp_file_path = tmp_file.name
        
        text = ""
        
        # Tentativa 1: Extrair texto diretamente com PyPDF2
        try:
            with open(tmp_file_path, "rb") as file:
                reader = PyPDF2.PdfReader(file)
                for page in reader.pages:
                    page_text = page.extract_text()
                    if page_text.strip():
                        text += page_text + "\n"
        except Exception as e:
            st.warning(f"PyPDF2 falhou: {e}")
        
        # Tentativa 2: Se não conseguiu extrair texto, usar OCR
        if not text.strip():
            st.info("📸 PDF parece ser uma imagem. Usando OCR para extrair texto...")
            
            # Progresso para OCR
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                # Converte PDF para imagens
                status_text.text("Convertendo PDF para imagens...")
                images = convert_from_path(tmp_file_path, dpi=300)
                
                total_pages = len(images)
                status_text.text(f"Processando {total_pages} página(s) com OCR...")
                
                for i, img in enumerate(images):
                    # Atualiza progresso
                    progress = (i + 1) / total_pages
                    progress_bar.progress(progress)
                    status_text.text(f"Processando página {i+1} de {total_pages}...")
                    
                    # Aplica OCR na imagem
                    page_text = pytesseract.image_to_string(img, lang='por')
                    text += page_text + "\n"
                
                progress_bar.progress(1.0)
                status_text.text("✅ OCR concluído!")
                
                # Limpa os elementos de progresso após um tempo
                import time
                time.sleep(1)
                progress_bar.empty()
                status_text.empty()
                
            except Exception as ocr_error:
                st.error(f"Erro no OCR: {ocr_error}")
                st.error("Certifique-se de que o Tesseract está instalado corretamente.")
        
        # Remove o arquivo temporário
        os.unlink(tmp_file_path)
        
        return text
        
    except Exception as e:
        st.error(f"Erro geral ao processar PDF: {e}")
        return ""

def extract_field(patterns, text):
    """
    Procura por uma lista de padrões regex e retorna o valor encontrado
    """
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            value = match.group(1).strip()
            # Remove quebras de linha e espaços extras
            value = re.sub(r'\s+', ' ', value)
            return value
    return "Não encontrado"

def parse_tokio_data(text):
    """
    Extrai dados específicos da apólice Tokio Marine
    """
    # Dados do cabeçalho/cliente
    dados_header = {
        "NOME DO CLIENTE": extract_field([
            r"Proprietário[:\s]*(.+?)(?:\n|$)",
            r"ROD TRANSPORTES LTDA"
        ], text),
        "CNPJ": extract_field([
            r"CNPJ[:\s]*(.+?)(?:\n|$)",
            r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})"
        ], text),
        "APÓLICE": extract_field([
            r"Nr Apólice[:\s]*(.+?)(?:\n|$)",
            r"Apólice Congenere[:\s]*(.+?)(?:\n|$)",
            r"(\d{8,})"
        ], text),
        "VIGÊNCIA": extract_field([
            r"Venc Apólice[:\s]*(.+?)(?:\n|$)",
            r"(\d{2}/\d{2}/\d{4})"
        ], text),
    }

    # Dados do veículo - adaptados para o formato do PDF enviado
    dados_veiculo = {
        "DESCRIÇÃO DO ITEM": extract_field([
            r"Descrição do Item[:\s-]*(.+?)(?:\n|$)",
            r"Produto Auto Frota"
        ], text),
        "CEP DE PERNOITE DO VEÍCULO": extract_field([
            r"CEP de Pernoite do Veículo[:\s]*(.+?)(?:\n|$)",
            r"(\d{5}-\d{3})"
        ], text),
        "TIPO DE UTILIZAÇÃO": extract_field([
            r"Tipo de utilização[:\s]*(.+?)(?:\n|$)",
            r"Particular/Comercial"
        ], text),
        "FABRICANTE": extract_field([
            r"Fabricante[:\s]*(.+?)(?:\n|$)",
            r"CHEVROLET"
        ], text),
        "VEÍCULO": extract_field([
            r"Veículo[:\s]*(.+?)(?:\n|$)",
            r"S10 PICK-UP LTZ.*"
        ], text),
        "ANO MODELO": extract_field([
            r"Ano Modelo[:\s]*(.+?)(?:\n|$)",
            r"(\d{4})"
        ], text),
        "CHASSI": extract_field([
            r"Chassi[:\s]*(.+?)(?:\n|$)",
            r"([A-Z0-9]{17})"
        ], text),
        "CHASSI REMARCADO": extract_field([
            r"Chassi Remarcado[:\s]*(.+?)(?:\n|$)"
        ], text),
        "PLACA": extract_field([
            r"Placa[:\s]*(.+?)(?:\n|$)",
            r"([A-Z]{3}\d{4}|[A-Z]{3}\d[A-Z]\d{2})"
        ], text),
        "COMBUSTÍVEL": extract_field([
            r"Combustível[:\s]*(.+?)(?:\n|$)",
            r"Diesel"
        ], text),
        "LOTAÇÃO VEÍCULO": extract_field([
            r"Lotação Veículo[:\s]*(.+?)(?:\n|$)",
            r"(\d+)"
        ], text),
        "VEÍCULO 0KM": extract_field([
            r"Veículo 0km[:\s]*(.+?)(?:\n|$)"
        ], text),
        "VEÍCULO BLINDADO": extract_field([
            r"Veículo Blindado[:\s]*(.+?)(?:\n|$)"
        ], text),
        "VEÍCULO COM KIT GÁS": extract_field([
            r"Veículo com Kit Gás[:\s]*(.+?)(?:\n|$)"
        ], text),
        "TIPO DE CARROCERIA": extract_field([
            r"Tipo de Carroceria[:\s]*(.+?)(?:\n|$)"
        ], text),
        "4º EIXO ADAPTADO": extract_field([
            r"4º Eixo Adaptado[:\s]*(.+?)(?:\n|$)"
        ], text),
        "CABINE SUPLEMENTAR": extract_field([
            r"Cabine Suplementar[:\s]*(.+?)(?:\n|$)"
        ], text),
        "DISPOSITIVO EM COMODATO": extract_field([
            r"Dispositivo em Comodato[:\s]*(.+?)(?:\n|$)"
        ], text),
        "ISENÇÃO FISCAL": extract_field([
            r"Isenção Fiscal[:\s]*(.+?)(?:\n|$)"
        ], text),
        "PROPRIETÁRIO": extract_field([
            r"Proprietário[:\s]*(.+?)(?:\n|$)",
            r"ROD TRANSPORTES LTDA"
        ], text),
        "FIPE": extract_field([
            r"Fipe[:\s]*(.+?)(?:\n|$)",
            r"(\d{6}-\d)"
        ], text),
        "TIPO DE SEGURO": extract_field([
            r"Tipo de Seguro[:\s]*(.+?)(?:\n|$)",
            r"Renovação Tokio sem sinistro"
        ], text),
        "NR APÓLICE CONGENERE": extract_field([
            r"Nr Apólice Congenere[:\s]*(.+?)(?:\n|$)",
            r"(\d{8})"
        ], text),
        "NOME DA CONGENERE": extract_field([
            r"Nome da Congenere[:\s]*(.+?)(?:\n|$)",
            r"TOKIO MARINE"
        ], text),
        "VENC APÓLICE CONGENERE": extract_field([
            r"Venc Apólice Cong[:\s\.]*(.+?)(?:\n|$)",
            r"(\d{2}/\d{2}/\d{4})"
        ], text),
    }

    return dados_header, dados_veiculo

def create_excel_file(dados_header, dados_veiculo):
    """
    Cria arquivo Excel com os dados extraídos
    """
    # Cria um buffer em memória
    buffer = BytesIO()
    
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        # Aba com dados gerais
        df_header = pd.DataFrame([dados_header])
        df_header.to_excel(writer, sheet_name='Dados Gerais', index=False)
        
        # Aba com dados do veículo
        df_veiculo = pd.DataFrame([dados_veiculo])
        df_veiculo.to_excel(writer, sheet_name='Veículos', index=False)
    
    buffer.seek(0)
    return buffer

def main():
    st.title("🚗 Conversor de Apólices Tokio Marine")
    st.markdown("---")
    
    st.markdown("""
    ### Como usar:
    1. 📤 Faça o upload da sua apólice em PDF (texto ou imagem)
    2. ⚡ Aguarde o processamento automático (OCR se necessário)
    3. 👀 Visualize os dados extraídos
    4. 💾 Baixe a planilha Excel gerada
    
    **✨ Suporte completo a PDFs digitalizados e escaneados!**
    """)
    
    # Upload do arquivo
    uploaded_file = st.file_uploader(
        "Escolha um arquivo PDF da apólice Tokio Marine",
        type=['pdf'],
        help="Faça upload do arquivo PDF da apólice para conversão (suporta PDFs com texto e imagens)"
    )
    
    if uploaded_file is not None:
        # Mostra informações do arquivo
        st.success(f"✅ Arquivo carregado: {uploaded_file.name}")
        st.info(f"📊 Tamanho: {len(uploaded_file.getvalue())/1024:.1f} KB")
        
        # Botão para processar
        if st.button("🔄 Processar PDF", type="primary"):
            # Extrai texto do PDF (com OCR se necessário)
            text = extract_text_from_pdf(uploaded_file)
            
            if text.strip():
                st.success("✅ Texto extraído com sucesso!")
                
                # Parse dos dados
                with st.spinner("Analisando dados da apólice..."):
                    dados_header, dados_veiculo = parse_tokio_data(text)
                
                # Mostra os dados extraídos
                st.markdown("## 📋 Dados Extraídos")
                
                # Dados gerais em duas colunas
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("### 🏢 Informações do Cliente")
                    for key, value in dados_header.items():
                        if value != "Não encontrado":
                            st.success(f"**{key}:** {value}")
                        else:
                            st.warning(f"**{key}:** {value}")
                
                with col2:
                    st.markdown("### 🚙 Informações do Veículo")
                    # Mostra apenas os campos encontrados
                    campos_importantes = [
                        "FABRICANTE", "VEÍCULO", "ANO MODELO", "PLACA", 
                        "CHASSI", "COMBUSTÍVEL", "FIPE", "PROPRIETÁRIO"
                    ]
                    for campo in campos_importantes:
                        if campo in dados_veiculo:
                            value = dados_veiculo[campo]
                            if value != "Não encontrado":
                                st.success(f"**{campo}:** {value}")
                            else:
                                st.warning(f"**{campo}:** {value}")
                
                # Tabelas expandidas
                with st.expander("📊 Ver todos os dados em tabela"):
                    st.markdown("#### Dados Gerais")
                    st.dataframe(pd.DataFrame([dados_header]), use_container_width=True)
                    
                    st.markdown("#### Dados do Veículo")
                    st.dataframe(pd.DataFrame([dados_veiculo]), use_container_width=True)
                
                # Gera o arquivo Excel
                excel_buffer = create_excel_file(dados_header, dados_veiculo)
                
                # Botão de download
                st.markdown("## 💾 Download")
                apolice_numero = dados_header.get('APÓLICE', 'sem_numero')
                if apolice_numero == "Não encontrado":
                    apolice_numero = "sem_numero"
                nome_arquivo = f"apolice_{apolice_numero}.xlsx"
                
                st.download_button(
                    label="📥 Baixar Planilha Excel",
                    data=excel_buffer,
                    file_name=nome_arquivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
                st.success("✅ Processamento concluído com sucesso!")
                
            else:
                st.error("❌ Não foi possível extrair texto do PDF. Verifique se o arquivo não está corrompido.")
        
        # Mostra preview do texto extraído (opcional)
        with st.expander("🔍 Ver texto extraído do PDF (debug)"):
            if st.button("Extrair texto para visualização"):
                with st.spinner("Extraindo texto..."):
                    text_preview = extract_text_from_pdf(uploaded_file)
                if text_preview:
                    st.text_area(
                        "Texto extraído:", 
                        text_preview[:3000] + "..." if len(text_preview) > 3000 else text_preview, 
                        height=400
                    )
                else:
                    st.error("Não foi possível extrair texto do PDF")

    # Informações adicionais
    with st.sidebar:
        st.markdown("## ℹ️ Informações")
        st.markdown("""
        **Recursos:**
        - 📄 PDFs com texto nativo
        - 📸 PDFs escaneados (OCR)
        - 🇧🇷 Reconhecimento em português
        - 📊 Export para Excel
        - 🔍 Modo debug
        
        **Requisitos:**
        - Tesseract OCR instalado
        - PDF da Tokio Marine
        """)
        
        st.markdown("## 🛠️ Configuração OCR")
        if st.button("Testar Tesseract"):
            try:
                version = pytesseract.get_tesseract_version()
                st.success(f"✅ Tesseract v{version}")
            except:
                st.error("❌ Tesseract não encontrado")
                st.markdown("""
                **Para instalar:**
                - Windows: [Download](https://github.com/UB-Mannheim/tesseract/wiki)
                - Linux: `sudo apt install tesseract-ocr tesseract-ocr-por`
                - Mac: `brew install tesseract tesseract-lang`
                """)

if __name__ == "__main__":
    main()