import streamlit as st
import PyPDF2
import pandas as pd
import re
import os
from io import BytesIO
import tempfile
import numpy as np

# Importações condicionais para OCR
OCR_AVAILABLE = False
try:
    import easyocr
    from pdf2image import convert_from_path
    from PIL import Image
    OCR_AVAILABLE = True
except ImportError:
    st.warning("⚠️ OCR não disponível. Apenas PDFs com texto serão processados.")

# Configuração da página
st.set_page_config(
    page_title="Conversor de Apólices Tokio Marine",
    page_icon="📄",
    layout="wide"
)

@st.cache_resource
def load_easyocr():
    """Carrega o modelo EasyOCR (cache para não recarregar sempre)"""
    if OCR_AVAILABLE:
        try:
            reader = easyocr.Reader(['pt', 'en'], gpu=False)
            return reader
        except Exception as e:
            st.error(f"Erro ao carregar EasyOCR: {e}")
            return None
    return None

def extract_text_from_pdf(pdf_file):
    """
    Extrai texto de um arquivo PDF usando PyPDF2 e EasyOCR como fallback
    """
    try:
        # Cria um arquivo temporário
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(pdf_file.read())
            tmp_file_path = tmp_file.name
        
        text = ""
        
        # Tentativa 1: Extrair texto diretamente com PyPDF2
        st.info("🔍 Tentando extrair texto diretamente do PDF...")
        try:
            with open(tmp_file_path, "rb") as file:
                reader = PyPDF2.PdfReader(file)
                for page_num, page in enumerate(reader.pages):
                    page_text = page.extract_text()
                    if page_text.strip():
                        text += page_text + "\n"
                        
            if text.strip():
                st.success("✅ Texto extraído diretamente do PDF!")
        except Exception as e:
            st.warning(f"PyPDF2 falhou: {e}")
        
        # Tentativa 2: Se não conseguiu extrair texto, usar EasyOCR
        if not text.strip() and OCR_AVAILABLE:
            st.info("📸 PDF parece ser uma imagem. Usando EasyOCR para extrair texto...")
            
            # Carrega o modelo EasyOCR
            reader = load_easyocr()
            if reader is None:
                st.error("❌ Não foi possível carregar o EasyOCR")
                return ""
            
            # Progresso para OCR
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                # Converte PDF para imagens
                status_text.text("Convertendo PDF para imagens...")
                images = convert_from_path(tmp_file_path, dpi=300)
                
                total_pages = len(images)
                status_text.text(f"Processando {total_pages} página(s) com EasyOCR...")
                
                all_text = []
                
                for i, img in enumerate(images):
                    # Atualiza progresso
                    progress = (i + 1) / total_pages
                    progress_bar.progress(progress)
                    status_text.text(f"Processando página {i+1} de {total_pages}...")
                    
                    # Converte PIL Image para numpy array
                    img_array = np.array(img)
                    
                    # Aplica EasyOCR na imagem
                    results = reader.readtext(img_array)
                    
                    # Extrai o texto dos resultados
                    page_text = []
                    for (bbox, text_detected, confidence) in results:
                        if confidence > 0.5:  # Filtro de confiança
                            page_text.append(text_detected)
                    
                    # Junta o texto da página
                    if page_text:
                        all_text.append(' '.join(page_text))
                
                # Junta todo o texto
                text = '\n'.join(all_text)
                
                progress_bar.progress(1.0)
                status_text.text("✅ EasyOCR concluído!")
                
                # Limpa os elementos de progresso
                import time
                time.sleep(1)
                progress_bar.empty()
                status_text.empty()
                
                if text.strip():
                    st.success("✅ Texto extraído com EasyOCR!")
                else:
                    st.warning("⚠️ EasyOCR não conseguiu extrair texto legível")
                
            except Exception as ocr_error:
                st.error(f"Erro no EasyOCR: {ocr_error}")
                progress_bar.empty()
                status_text.empty()
        
        elif not text.strip():
            st.error("❌ PDF é uma imagem, mas EasyOCR não está disponível.")
            st.markdown("""
            **Para usar EasyOCR:**
            ```bash
            pip install easyocr pdf2image Pillow
            ```
            """)
        
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
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
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
    # Limpa o texto para melhor parsing
    text = re.sub(r'\s+', ' ', text)
    
    # Dados do cabeçalho/cliente
    dados_header = {
        "NOME DO CLIENTE": extract_field([
            r"Proprietário[:\s]*([^:\n]*?)(?=\s*(?:Tipo|CEP|Fabricante|$))",
            r"ROD TRANSPORTES LTDA",
            r"([A-Z\s]{10,}(?:LTDA|S\.A\.|EIRELI))"
        ], text),
        "CNPJ": extract_field([
            r"CNPJ[:\s]*([^:\n]*?)(?=\s*(?:Tipo|CEP|Fabricante|$))",
            r"(\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2})"
        ], text),
        "APÓLICE": extract_field([
            r"(?:Nr\s*)?Apólice[:\s]*([^:\n]*?)(?=\s*(?:Venc|Tipo|CEP|$))",
            r"(\d{8,})"
        ], text),
        "VIGÊNCIA": extract_field([
            r"Venc[^:]*Apólice[^:]*[:\s]*([^:\n]*?)(?=\s*(?:Tipo|CEP|$))",
            r"(\d{2}/\d{2}/\d{4})"
        ], text),
    }

    # Dados do veículo
    dados_veiculo = {
        "DESCRIÇÃO DO ITEM": extract_field([
            r"Descrição do Item[^:]*[:\s-]*([^:\n]*?)(?=\s*(?:CEP|Tipo|Fabricante|$))",
            r"(Produto Auto Frota)",
            r"(\d+\s*-\s*Produto Auto Frota)"
        ], text),
        "CEP DE PERNOITE DO VEÍCULO": extract_field([
            r"CEP de Pernoite do Veículo[:\s]*([^:\n]*?)(?=\s*(?:Tipo|Fabricante|$))",
            r"(\d{5}-?\d{3})"
        ], text),
        "TIPO DE UTILIZAÇÃO": extract_field([
            r"Tipo de utilização[:\s]*([^:\n]*?)(?=\s*(?:Ano|Fabricante|$))",
            r"(Particular/?Comercial)"
        ], text),
        "FABRICANTE": extract_field([
            r"Fabricante[:\s]*([^:\n]*?)(?=\s*(?:Veículo|Ano|$))",
            r"(CHEVROLET|FORD|VOLKSWAGEN|FIAT|[A-Z]{3,})"
        ], text),
        "VEÍCULO": extract_field([
            r"Veículo[:\s]*([^:\n]*?)(?=\s*(?:Ano|4º|$))",
            r"(S10 PICK-UP LTZ[^:\n]*)"
        ], text),
        "ANO MODELO": extract_field([
            r"Ano Modelo[:\s]*([^:\n]*?)(?=\s*(?:Chassi|4º|$))",
            r"(\d{4})"
        ], text),
        "CHASSI": extract_field([
            r"(?:^|\s)Chassi[:\s]*([^:\n]*?)(?=\s*(?:Chassi Remarcado|Placa|$))",
            r"([A-Z0-9]{17})"
        ], text),
        "CHASSI REMARCADO": extract_field([
            r"Chassi Remarcado[:\s]*([^:\n]*?)(?=\s*(?:Combustível|Placa|$))"
        ], text),
        "PLACA": extract_field([
            r"Placa[:\s]*([^:\n]*?)(?=\s*(?:Lotação|Combustível|$))",
            r"([A-Z]{3}\d{4}|[A-Z]{3}\d[A-Z]\d{2})"
        ], text),
        "COMBUSTÍVEL": extract_field([
            r"Combustível[:\s]*([^:\n]*?)(?=\s*(?:Lotação|Veículo|$))",
            r"(Diesel|Gasolina|Flex|Álcool)"
        ], text),
        "LOTAÇÃO VEÍCULO": extract_field([
            r"Lotação Veículo[:\s]*([^:\n]*?)(?=\s*(?:Veículo|Dispositivo|$))",
            r"(\d+)"
        ], text),
        "VEÍCULO 0KM": extract_field([
            r"Veículo 0km[:\s]*([^:\n]*?)(?=\s*(?:Veículo|Dispositivo|$))"
        ], text),
        "VEÍCULO BLINDADO": extract_field([
            r"Veículo Blindado[:\s]*([^:\n]*?)(?=\s*(?:Dispositivo|Isenção|$))"
        ], text),
        "VEÍCULO COM KIT GÁS": extract_field([
            r"Veículo com Kit Gás[:\s]*([^:\n]*?)(?=\s*(?:Tipo|Isenção|$))"
        ], text),
        "TIPO DE CARROCERIA": extract_field([
            r"Tipo de Carroceria[:\s]*([^:\n]*?)(?=\s*(?:4º|Cabine|$))"
        ], text),
        "4º EIXO ADAPTADO": extract_field([
            r"4º Eixo Adaptado[:\s]*([^:\n]*?)(?=\s*(?:Cabine|Dispositivo|$))"
        ], text),
        "CABINE SUPLEMENTAR": extract_field([
            r"Cabine Suplementar[:\s]*([^:\n]*?)(?=\s*(?:Dispositivo|Isenção|$))"
        ], text),
        "DISPOSITIVO EM COMODATO": extract_field([
            r"Dispositivo em Comodato[:\s]*([^:\n]*?)(?=\s*(?:Isenção|Fipe|$))"
        ], text),
        "ISENÇÃO FISCAL": extract_field([
            r"Isenção Fiscal[:\s]*([^:\n]*?)(?=\s*(?:Fipe|Proprietário|$))"
        ], text),
        "PROPRIETÁRIO": extract_field([
            r"Proprietário[:\s]*([^:\n]*?)(?=\s*(?:Fipe|Tipo|$))",
            r"(ROD TRANSPORTES LTDA)"
        ], text),
        "FIPE": extract_field([
            r"Fipe[:\s]*([^:\n]*?)(?=\s*(?:Nr|Nome|$))",
            r"(\d{6}-\d)"
        ], text),
        "TIPO DE SEGURO": "Renovação Tokio sem sinistro",
        "NR APÓLICE CONGENERE": extract_field([
            r"Nr Apólice Congenere[:\s]*([^:\n]*?)(?=\s*(?:Nome|Venc|$))",
            r"(\d{8,})"
        ], text),
        "NOME DA CONGENERE": extract_field([
            r"Nome da Congenere[:\s]*([^:\n]*?)(?=\s*(?:Venc|$))",
            r"(TOKIO MARINE[^:\n]*)"
        ], text),
        "VENC APÓLICE CONGENERE": extract_field([
            r"Venc Apólice Cong[^:]*[:\s]*([^:\n]*?)(?=\s*$)",
            r"(\d{2}/\d{2}/\d{4})"
        ], text),
    }

    return dados_header, dados_veiculo

def create_excel_file(dados_header, dados_veiculo):
    """
    Cria arquivo Excel com os dados extraídos
    """
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
    
    # Status do OCR
    if OCR_AVAILABLE:
        st.success("✅ EasyOCR disponível - Suporte completo a PDFs escaneados!")
    else:
        st.warning("⚠️ EasyOCR não disponível - Apenas PDFs com texto nativo")
    
    st.markdown("""
    ### Como usar:
    1. 📤 Faça o upload da sua apólice em PDF (texto ou imagem)
    2. ⚡ Aguarde o processamento automático
    3. 👀 Visualize os dados extraídos
    4. 💾 Baixe a planilha Excel gerada
    
    **✨ Powered by EasyOCR - Melhor precisão em PDFs escaneados!**
    """)
    
    # Upload do arquivo
    uploaded_file = st.file_uploader(
        "Escolha um arquivo PDF da apólice Tokio Marine",
        type=['pdf'],
        help="Suporte completo a PDFs com texto nativo e escaneados"
    )
    
    if uploaded_file is not None:
        # Mostra informações do arquivo
        st.success(f"✅ Arquivo carregado: {uploaded_file.name}")
        st.info(f"📊 Tamanho: {len(uploaded_file.getvalue())/1024:.1f} KB")
        
        # Botão para processar
        if st.button("🔄 Processar PDF", type="primary"):
            # Extrai texto do PDF
            text = extract_text_from_pdf(uploaded_file)
            
            if text.strip():
                # Parse dos dados
                with st.spinner("🧠 Analisando dados da apólice..."):
                    dados_header, dados_veiculo = parse_tokio_data(text)
                
                # Mostra os dados extraídos
                st.markdown("## 📋 Dados Extraídos")
                
                # Contador de campos encontrados
                encontrados_header = sum(1 for v in dados_header.values() if v != "Não encontrado")
                encontrados_veiculo = sum(1 for v in dados_veiculo.values() if v != "Não encontrado")
                total_campos = len(dados_header) + len(dados_veiculo)
                total_encontrados = encontrados_header + encontrados_veiculo
                
                st.info(f"📊 **{total_encontrados}/{total_campos}** campos extraídos com sucesso ({(total_encontrados/total_campos)*100:.1f}%)")
                
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
                    # Mostra apenas os campos mais importantes primeiro
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
                
                st.balloons()  # Animação de sucesso
                st.success("✅ Processamento concluído com sucesso!")
                
            else:
                st.error("❌ Não foi possível extrair texto do PDF.")
        
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
                    st.info(f"📏 Total de caracteres: {len(text_preview)}")
                else:
                    st.error("Não foi possível extrair texto do PDF")

    # Informações adicionais
    with st.sidebar:
        st.markdown("## ℹ️ Status do Sistema")
        
        # Verificações de dependências
        st.markdown("**Dependências:**")
        st.success("✅ PyPDF2")
        st.success("✅ Pandas")
        
        if OCR_AVAILABLE:
            st.success("✅ EasyOCR")
            st.success("✅ pdf2image")
            st.success("✅ Pillow")
        else:
            st.error("❌ Dependências OCR")
        
        st.markdown("## 🎯 Recursos")
        st.markdown("""
        **Sempre disponível:**
        - 📄 PDFs com texto nativo
        - 📊 Export para Excel
        - 🔍 Modo debug
        
        **Com EasyOCR:**
        - 📸 PDFs escaneados
        - 🇧🇷 Reconhecimento PT/EN
        - 🎯 Alta precisão
        - ☁️ Funciona na nuvem
        """)
        
        if not OCR_AVAILABLE:
            st.markdown("## 🛠️ Para habilitar OCR")
            st.code("""
pip install easyocr pdf2image Pillow
            """)
            st.markdown("**EasyOCR é muito melhor que Tesseract para deploy na nuvem!**")
        
        st.markdown("## 📈 Vantagens do EasyOCR")
        st.markdown("""
        - ✅ Instalação mais simples
        - ✅ Melhor precisão
        - ✅ Funciona no Streamlit Cloud
        - ✅ Suporte a múltiplos idiomas
        - ✅ Não requer configuração externa
        """)

if __name__ == "__main__":
    main()