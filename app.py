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
TESSERACT_AVAILABLE = False

try:
    import pytesseract
    from pdf2image import convert_from_path
    from PIL import Image
    TESSERACT_AVAILABLE = True
    OCR_AVAILABLE = True
except ImportError:
    try:
        import easyocr
        from pdf2image import convert_from_path
        from PIL import Image
        import numpy as np
        OCR_AVAILABLE = True
    except ImportError:
        pass

# Configuração da página
st.set_page_config(
    page_title="Conversor de Apólices Tokio Marine",
    page_icon="📄",
    layout="wide"
)

@st.cache_resource
def load_easyocr():
    """Carrega o modelo EasyOCR apenas se necessário"""
    if not TESSERACT_AVAILABLE and OCR_AVAILABLE:
        try:
            import easyocr
            reader = easyocr.Reader(['pt'], gpu=False, verbose=False)
            return reader
        except Exception as e:
            st.error(f"Erro ao carregar EasyOCR: {e}")
            return None
    return None

def extract_text_from_pdf(pdf_file):
    """Extrai texto de um arquivo PDF usando PyPDF2 e OCR como fallback"""
    try:
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
        
        # Tentativa 2: OCR apenas se necessário
        if not text.strip() and OCR_AVAILABLE:
            if TESSERACT_AVAILABLE:
                st.info("📸 PDF é imagem. Usando Tesseract OCR...")
                try:
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    status_text.text("Convertendo PDF para imagens...")
                    images = convert_from_path(tmp_file_path, dpi=200)
                    
                    total_pages = len(images)
                    all_text = []
                    
                    for i, img in enumerate(images):
                        progress = (i + 1) / total_pages
                        progress_bar.progress(progress)
                        status_text.text(f"OCR página {i+1} de {total_pages}...")
                        
                        page_text = pytesseract.image_to_string(img, lang='por')
                        if page_text.strip():
                            all_text.append(page_text)
                    
                    text = '\n'.join(all_text)
                    progress_bar.progress(1.0)
                    status_text.text("✅ Tesseract OCR concluído!")
                    
                    import time
                    time.sleep(1)
                    progress_bar.empty()
                    status_text.empty()
                    
                except Exception as tesseract_error:
                    st.warning(f"Tesseract falhou: {tesseract_error}")
                    text = ""
            
            if not text.strip() and not TESSERACT_AVAILABLE:
                st.warning("⚠️ Usando EasyOCR - Pode demorar no primeiro uso...")
                
                download_warning = st.warning("📥 EasyOCR está baixando modelos. Isso pode demorar alguns minutos na primeira vez...")
                
                reader = load_easyocr()
                if reader:
                    try:
                        download_warning.empty()
                        
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        status_text.text("Convertendo PDF para imagens...")
                        images = convert_from_path(tmp_file_path, dpi=200)
                        
                        total_pages = len(images)
                        all_text = []
                        
                        for i, img in enumerate(images):
                            progress = (i + 1) / total_pages
                            progress_bar.progress(progress)
                            status_text.text(f"EasyOCR página {i+1} de {total_pages}...")
                            
                            img_array = np.array(img)
                            results = reader.readtext(img_array)
                            
                            page_text = []
                            for (bbox, text_detected, confidence) in results:
                                if confidence > 0.6:
                                    page_text.append(text_detected)
                            
                            if page_text:
                                all_text.append(' '.join(page_text))
                        
                        text = '\n'.join(all_text)
                        progress_bar.progress(1.0)
                        status_text.text("✅ EasyOCR concluído!")
                        
                        import time
                        time.sleep(1)
                        progress_bar.empty()
                        status_text.empty()
                        
                    except Exception as easyocr_error:
                        st.error(f"EasyOCR falhou: {easyocr_error}")
                        download_warning.empty()
        
        os.unlink(tmp_file_path)
        return text
        
    except Exception as e:
        st.error(f"Erro geral ao processar PDF: {e}")
        return ""

def extract_field(patterns, text):
    """Procura por uma lista de padrões regex e retorna o valor encontrado"""
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            if match.groups():
                value = match.group(1).strip()
                value = re.sub(r'\s+', ' ', value)
                return value
            else:
                value = match.group(0).strip()
                value = re.sub(r'\s+', ' ', value)
                return value
    return ""

def extract_vehicle_sections(text):
    """Extrai todas as seções de veículos do texto"""
    # Padrão para encontrar seções de veículos
    vehicle_pattern = r'Descrição do Item -\s*(\d+)\s*-\s*Produto Auto Frota(.*?)(?=Descrição do Item -\s*\d+\s*-\s*Produto Auto Frota|$)'
    
    sections = re.findall(vehicle_pattern, text, re.DOTALL | re.IGNORECASE)
    
    vehicles = []
    for item_num, content in sections:
        vehicles.append({
            'item': item_num.strip(),
            'content': content.strip()
        })
    
    return vehicles

def parse_vehicle_data(vehicle_content, item_num):
    """Extrai dados de um veículo específico"""
    
    # Função auxiliar para extrair valores monetários
    def extract_money(patterns, text):
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                value = match.group(1).replace('.', '').replace(',', '.')
                try:
                    return float(value)
                except:
                    return value
        return 0
    
    # Função auxiliar para extrair valores simples
    def extract_simple(patterns, text):
        result = extract_field(patterns, text)
        return result if result else ""
    
    dados = {
        # Identificação Básica
        'Item': item_num,
        'CEP de Pernoite': extract_simple([
            r'CEP de Pernoite do Veículo[:\s]*([^\s]+)',
            r'(\d{5}-?\d{3})'
        ], vehicle_content),
        
        'Fabricante': extract_simple([
            r'Fabricante[:\s]*([^\n\r]+?)(?=\s*Veículo:|$)',
            r'(CHEVROLET|FORD|VOLKSWAGEN|FIAT|NISSAN|TOYOTA|BYD|MITSUBISHI)'
        ], vehicle_content),
        
        'Veículo': extract_simple([
            r'Veículo[:\s]*([^\n\r]+?)(?=\s*(?:Ano Modelo|4º Eixo)|$)'
        ], vehicle_content),
        
        'Ano Modelo': extract_simple([
            r'Ano Modelo[:\s]*(\d{4})'
        ], vehicle_content),
        
        'Chassi': extract_simple([
            r'Chassi[:\s]*([A-Z0-9]{17})'
        ], vehicle_content),
        
        'Chassi Remarcado': extract_simple([
            r'Chassi Remarcado[:\s]*([^\n\r]+?)(?=\s*Placa:|$)'
        ], vehicle_content),
        
        'Placa': extract_simple([
            r'Placa[:\s]*([A-Z0-9]+)'
        ], vehicle_content),
        
        'Combustível': extract_simple([
            r'Combustível[:\s]*([^\n\r]+?)(?=\s*Lotação|$)',
            r'(Diesel|Gasolina|Flex|Álcool|Elétrico)'
        ], vehicle_content),
        
        'Lotação Veículo': extract_simple([
            r'Lotação Veículo[:\s]*(\d+)'
        ], vehicle_content),
        
        'Veículo 0km': extract_simple([
            r'Veículo 0km[:\s]*([^\n\r]+?)(?=\s*Veículo Blindado|$)'
        ], vehicle_content),
        
        'Veículo Blindado': extract_simple([
            r'Veículo Blindado[:\s]*([^\n\r]+?)(?=\s*Veículo com Kit|$)'
        ], vehicle_content),
        
        'Veículo com Kit Gás': extract_simple([
            r'Veículo com Kit Gás[:\s]*([^\n\r]+?)(?=\s*Dispositivo|$)'
        ], vehicle_content),
        
        'Dispositivo em Comodato': extract_simple([
            r'Dispositivo em Comodato[:\s]*([^\n\r]+?)(?=\s*Tipo de|$)'
        ], vehicle_content),
        
        'Tipo de Carroceria': extract_simple([
            r'Tipo de Carroceria[:\s]*([^\n\r]+?)(?=\s*Isenção|$)'
        ], vehicle_content),
        
        'Isenção Fiscal': extract_simple([
            r'Isenção Fiscal[:\s]*([^\n\r]+?)(?=\s*Proprietário|$)'
        ], vehicle_content),
        
        'Proprietário': extract_simple([
            r'Proprietário[:\s]*([^\n\r]+?)(?=\s*Fipe|$)'
        ], vehicle_content),
        
        'Fipe': extract_simple([
            r'Fipe[:\s]*([^\n\r]+?)(?=\s*Tipo de Seguro|$)'
        ], vehicle_content),
        
        'Tipo de Seguro': extract_simple([
            r'Tipo de Seguro[:\s]*([^\n\r]+?)(?=\s*Nr Apólice|$)'
        ], vehicle_content),
        
        'Nome da Seguradora Anterior': extract_simple([
            r'Nome da Congenere[:\s]*([^\n\r]+?)(?=\s*Venc Apólice|$)'
        ], vehicle_content),
        
        'Nr Apólice Congênere': extract_simple([
            r'Nr Apólice Congenere[:\s]*([^\n\r]+?)(?=\s*Venc|$)'
        ], vehicle_content),
        
        'Venc Apólice Cong.': extract_simple([
            r'Venc Apólice Cong\.[:\s]*([^\n\r]+?)(?=\s*Classe|$)'
        ], vehicle_content),
        
        'Classe de Bônus': extract_simple([
            r'Classe de Bônus[:\s]*(\d+)'
        ], vehicle_content),
        
        'Código de Identificação (CI)': extract_simple([
            r'Código de Identificação \(CI\)[:\s]*([^\n\r]+?)(?=\s*Km de|$)'
        ], vehicle_content),
        
        'Km de Reboque': extract_simple([
            r'Km de Reboque[:\s]*([^\n\r]+?)(?=\s*CNPJ|$)'
        ], vehicle_content),
        
        'CNPJ Fornecedor': extract_simple([
            r'CNPJ[:\s]*([^\n\r]+?)(?=\s*Fornecedor|$)'
        ], vehicle_content),
        
        'Fornecedor de Vidros': extract_simple([
            r'Fornecedor de Vidros[:\s]*([^\n\r]+?)(?=\s*Coberturas|$)'
        ], vehicle_content),
    }
    
    # Extração de valores das coberturas
    # Colisão, Incêndio e Roubo/Furto
    dados['Limite Colisão/Incêndio/Roubo'] = extract_simple([
        r'Colisão, Incêndio e Roubo/Furto\s+Valor Referenciado \(VMR\)\s+([\d.,]+)'
    ], vehicle_content)
    
    dados['Prêmio Colisão/Incêndio/Roubo'] = extract_money([
        r'Colisão, Incêndio e Roubo/Furto.*?Valor Referenciado.*?\s+([\d.,]+)\s+([\d.,]+)'
    ], vehicle_content)
    
    dados['Franquia Colisão/Incêndio/Roubo'] = extract_money([
        r'Colisão, Incêndio e Roubo/Furto.*?Valor Referenciado.*?\s+[\d.,]+\s+[\d.,]+\s+([\d.,]+)'
    ], vehicle_content)
    
    # RCF-V Danos Materiais
    dados['Limite RCF-V Danos Materiais'] = extract_money([
        r'RCF-V - Danos Materiais\s+([\d.,]+)'
    ], vehicle_content)
    
    dados['Prêmio RCF-V Danos Materiais'] = extract_money([
        r'RCF-V - Danos Materiais\s+[\d.,]+\s+([\d.,]+)'
    ], vehicle_content)
    
    # RCF-V Danos Corporais
    dados['Limite RCF-V Danos Corporais'] = extract_money([
        r'RCF-V - Danos Corporais\s+([\d.,]+)'
    ], vehicle_content)
    
    dados['Prêmio RCF-V Danos Corporais'] = extract_money([
        r'RCF-V - Danos Corporais\s+[\d.,]+\s+([\d.,]+)'
    ], vehicle_content)
    
    # RCF-V Danos Morais
    dados['Limite RCF-V Danos Morais'] = extract_money([
        r'RCF-V - Danos morais\s+([\d.,]+)'
    ], vehicle_content)
    
    dados['Prêmio RCF-V Danos Morais'] = extract_money([
        r'RCF-V - Danos morais\s+[\d.,]+\s+([\d.,]+)'
    ], vehicle_content)
    
    # APP Morte por Passageiro
    dados['Limite APP Morte por Passageiro'] = extract_money([
        r'APP - Morte por Passageiro\s+([\d.,]+)'
    ], vehicle_content)
    
    dados['Prêmio APP Morte por Passageiro'] = extract_money([
        r'APP - Morte por Passageiro\s+[\d.,]+\s+([\d.,]+)'
    ], vehicle_content)
    
    # APP Invalidez por Passageiro
    dados['Limite APP Invalidez por Passageiro'] = extract_money([
        r'APP - Invalidez por Passageiro\s+([\d.,]+)'
    ], vehicle_content)
    
    dados['Prêmio APP Invalidez por Passageiro'] = extract_money([
        r'APP - Invalidez por Passageiro\s+[\d.,]+\s+([\d.,]+)'
    ], vehicle_content)
    
    # Assistência 24 horas
    dados['Assistência 24 horas'] = extract_simple([
        r'Assistência 24 horas\s+([^\n\r]+?)(?=\s*Km adicional|$)'
    ], vehicle_content)
    
    dados['Prêmio Assistência 24h'] = extract_money([
        r'Assistência 24 horas\s+VIP\s+([\d.,]+)'
    ], vehicle_content)
    
    # Km adicional de reboque
    dados['Km adicional de reboque'] = extract_simple([
        r'Km adicional de reboque\s+([^\n\r]+?)(?=\s*[\d.,]+|$)'
    ], vehicle_content)
    
    dados['Prêmio Km adicional'] = extract_money([
        r'Km adicional de reboque\s+[\w\s]+\s+([\d.,]+)'
    ], vehicle_content)
    
    # Kit Gás
    dados['Valor Kit Gás'] = extract_money([
        r'Kit [gG]ás\s+([\d.,]+)\s+([\d.,]+)'
    ], vehicle_content)
    
    dados['Prêmio Kit Gás'] = extract_money([
        r'Kit [gG]ás\s+[\d.,]+\s+([\d.,]+)'
    ], vehicle_content)
    
    # Blindagem
    dados['Valor Blindagem'] = extract_money([
        r'Blindagem\s+([\d.,]+)\s+([\d.,]+)'
    ], vehicle_content)
    
    dados['Prêmio Blindagem'] = extract_money([
        r'Blindagem\s+[\d.,]+\s+([\d.,]+)'
    ], vehicle_content)
    
    # Prêmio Líquido Total
    dados['Prêmio Líquido Total'] = extract_money([
        r'Prêmio Líquido Total[:\s]+([\d.,]+)'
    ], vehicle_content)
    
    # Franquias de Vidros
    franquias_patterns = {
        'Franquia Parabrisa': r'Parabrisa[:\s]*R\$\s*([\d.,]+)',
        'Franquia Parabrisa Delaminado': r'Parabrisa Delaminado[:\s]*R\$\s*([\d.,]+)',
        'Franquia Vigia/Traseiro': r'Vigia/Traseiro[:\s]*R\$\s*([\d.,]+)',
        'Franquia Vigia/Traseiro Delaminado': r'Vigia/Traseiro Delaminado[:\s]*R\$\s*([\d.,]+)',
        'Franquia Lateral': r'Lateral[:\s]*R\$\s*([\d.,]+)',
        'Franquia Lateral Delaminado': r'Lateral Delaminado[:\s]*R\$\s*([\d.,]+)',
        'Franquia Farol Halógeno': r'Farol Halógeno[:\s]*R\$\s*([\d.,]+)',
        'Franquia Farol Xenon/LED': r'Farol xenon/led[:\s]*R\$\s*([\d.,]+)',
        'Franquia Farol Inteligente': r'Farol Inteligente[:\s]*R\$\s*([\d.,]+)',
        'Franquia Farol Auxiliar': r'Farol auxiliar[:\s]*R\$\s*([\d.,]+)',
        'Franquia Lanterna Halógena': r'Lanterna Halógena[:\s]*R\$\s*([\d.,]+)',
        'Franquia Lanterna LED': r'Lanterna led[:\s]*R\$\s*([\d.,]+)',
        'Franquia Lanterna Auxiliar': r'Lanterna auxiliar[:\s]*R\$\s*([\d.,]+)',
        'Franquia Retrovisor Externo': r'Retrovisor externo[:\s]*R\$\s*([\d.,]+)',
        'Franquia Retrovisor Interno': r'Retrovisor Interno[:\s]*R\$\s*([\d.,]+)',
        'Franquia Teto Solar': r'Teto Solar[:\s]*R\$\s*([\d.,]+)',
        'Franquia Máquina de Vidro': r'Máquina de Vidro[:\s]*R\$\s*([\d.,]+)'
    }
    
    for field, pattern in franquias_patterns.items():
        dados[field] = extract_money([pattern], vehicle_content)
    
    return dados

def parse_header_data(text):
    """Extrai dados gerais da apólice"""
    dados_header = {
        'Razão Social': extract_field([
            r'Razão Social[:\s]*([^\n\r]+?)(?=\s*CNPJ|$)',
            r'(ROD TRANSPORTES LTDA)'
        ], text),
        
        'CNPJ': extract_field([
            r'CNPJ[:\s]*([^\n\r]+?)(?=\s*Atividade|$)',
            r'(\d{3}\.\d{3}\.\d{3}/\d{4}-\d{2})'
        ], text),
        
        'Atividade Principal': extract_field([
            r'Atividade Principal[:\s]*([^\n\r]+?)(?=\s*Endereço|$)'
        ], text),
        
        'Endereço': extract_field([
            r'Endereço[:\s]*([^\n\r]+?)(?=\s*Bairro|$)'
        ], text),
        
        'Bairro': extract_field([
            r'Bairro[:\s]*([^\n\r]+?)(?=\s*CEP|$)'
        ], text),
        
        'CEP': extract_field([
            r'CEP[:\s]*([^\n\r]+?)(?=\s*Cidade|$)'
        ], text),
        
        'Cidade': extract_field([
            r'Cidade[:\s]*([^\n\r]+?)(?=\s*UF|$)'
        ], text),
        
        'UF': extract_field([
            r'UF[:\s]*([^\n\r]+?)(?=\s*Telefone|$)'
        ], text),
        
        'Telefone': extract_field([
            r'Telefone[:\s]*([^\n\r]+?)(?=\s*Celular|$)'
        ], text),
        
        'Celular': extract_field([
            r'Celular[:\s]*([^\n\r]+?)(?=\s*Dados|$)'
        ], text),
        
        'Ramo': extract_field([
            r'Ramo[:\s]*([^\n\r]+?)(?=\s*Apólice|$)'
        ], text),
        
        'Apólice': extract_field([
            r'Apólice[:\s]*([^\n\r]+?)(?=\s*Negócio|$)'
        ], text),
        
        'Negócio': extract_field([
            r'Negócio[:\s]*([^\n\r]+?)(?=\s*Proposta|$)'
        ], text),
        
        'Proposta': extract_field([
            r'Proposta[:\s]*([^\n\r]+?)(?=\s*Quantidade|$)'
        ], text),
        
        'Quantidade de Itens': extract_field([
            r'Quantidade de Itens[:\s]*([^\n\r]+?)(?=\s*Sucursal|$)'
        ], text),
        
        'Sucursal': extract_field([
            r'Sucursal[:\s]*([^\n\r]+?)(?=\s*Moeda|$)'
        ], text),
        
        'Moeda': extract_field([
            r'Moeda[:\s]*([^\n\r]+?)(?=\s*Forma|$)'
        ], text),
        
        'Vigência do Seguro': extract_field([
            r'Vigência do Seguro[:\s]*([^\n\r]+?)(?=\s*Data|$)'
        ], text),
        
        'Data da Versão': extract_field([
            r'Data da Versão[:\s]*([^\n\r]+?)(?=\s*Data da Emissão|$)'
        ], text),
        
        'Data da Emissão': extract_field([
            r'Data da Emissão[:\s]*([^\n\r]+?)(?=\s*Segurado|$)'
        ], text),
        
        'Nome Corretor': extract_field([
            r'Nome Corretor[:\s]*([^\n\r]+?)(?=\s*Part|$)'
        ], text),
        
        'Registro SUSEP': extract_field([
            r'Registro SUSEP[:\s]*([^\n\r]+?)(?=\s*Líder|$)'
        ], text),
        
        'Prêmio Líquido Total Geral': extract_field([
            r'Prêmio Líquido Total[:\s]*R\$\s*([^\n\r]+?)(?=\s*Juros|$)'
        ], text),
        
        'Juros': extract_field([
            r'Juros[:\s]*R\$\s*([^\n\r]+?)(?=\s*I\.O\.F|$)'
        ], text),
        
        'I.O.F': extract_field([
            r'I\.O\.F[:\s]*R\$\s*([^\n\r]+?)(?=\s*Prêmio Total|$)'
        ], text),
        
        'Prêmio Total Geral': extract_field([
            r'Prêmio Total[:\s]*R\$\s*([^\n\r]+?)(?=\s*Cobrança|$)'
        ], text),
        
        'Cobrança': extract_field([
            r'Cobrança[:\s]*([^\n\r]+?)(?=\s*Parcelamento|$)'
        ], text),
    }
    
    return dados_header

def create_excel_file(dados_header, all_vehicles_data):
    """Cria arquivo Excel com os dados extraídos"""
    buffer = BytesIO()
    
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        # Aba com dados gerais da apólice
        df_header = pd.DataFrame([dados_header])
        df_header.to_excel(writer, sheet_name='Dados Gerais', index=False)
        
        # Aba com todos os veículos
        if all_vehicles_data:
            df_vehicles = pd.DataFrame(all_vehicles_data)
            df_vehicles.to_excel(writer, sheet_name='Todos os Veículos', index=False)
            
            # Aba separada para dados financeiros (coberturas e prêmios)
            financial_columns = [
                'Item', 'Fabricante', 'Veículo', 'Placa',
                'Limite Colisão/Incêndio/Roubo', 'Prêmio Colisão/Incêndio/Roubo', 'Franquia Colisão/Incêndio/Roubo',
                'Limite RCF-V Danos Materiais', 'Prêmio RCF-V Danos Materiais',
                'Limite RCF-V Danos Corporais', 'Prêmio RCF-V Danos Corporais',
                'Limite RCF-V Danos Morais', 'Prêmio RCF-V Danos Morais',
                'Limite APP Morte por Passageiro', 'Prêmio APP Morte por Passageiro',
                'Limite APP Invalidez por Passageiro', 'Prêmio APP Invalidez por Passageiro',
                'Prêmio Assistência 24h', 'Prêmio Km adicional',
                'Valor Kit Gás', 'Prêmio Kit Gás',
                'Valor Blindagem', 'Prêmio Blindagem',
                'Prêmio Líquido Total'
            ]
            
            financial_data = []
            for vehicle in all_vehicles_data:
                financial_row = {col: vehicle.get(col, '') for col in financial_columns}
                financial_data.append(financial_row)
            
            df_financial = pd.DataFrame(financial_data)
            df_financial.to_excel(writer, sheet_name='Dados Financeiros', index=False)
            
            # Aba para franquias de vidros
            franquia_columns = [
                'Item', 'Fabricante', 'Veículo', 'Placa',
                'Franquia Parabrisa', 'Franquia Parabrisa Delaminado',
                'Franquia Vigia/Traseiro', 'Franquia Vigia/Traseiro Delaminado',
                'Franquia Lateral', 'Franquia Lateral Delaminado',
                'Franquia Farol Halógeno', 'Franquia Farol Xenon/LED', 'Franquia Farol Inteligente', 'Franquia Farol Auxiliar',
                'Franquia Lanterna Halógena', 'Franquia Lanterna LED', 'Franquia Lanterna Auxiliar',
                'Franquia Retrovisor Externo', 'Franquia Retrovisor Interno',
                'Franquia Teto Solar', 'Franquia Máquina de Vidro'
            ]
            
            franquia_data = []
            for vehicle in all_vehicles_data:
                franquia_row = {col: vehicle.get(col, '') for col in franquia_columns}
                franquia_data.append(franquia_row)
            
            df_franquias = pd.DataFrame(franquia_data)
            df_franquias.to_excel(writer, sheet_name='Franquias Vidros', index=False)
    
    buffer.seek(0)
    return buffer

def main():
    st.title("🚗 Conversor Completo de Apólices Tokio Marine")
    st.markdown("---")
    
    # Status detalhado do sistema
    st.markdown("### 🔧 Status do Sistema")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.success("✅ PyPDF2 (Texto)")
    with col2:
        if TESSERACT_AVAILABLE:
            st.success("✅ Tesseract (OCR)")
        elif OCR_AVAILABLE:
            st.warning("⚠️ EasyOCR (Lento)")
        else:
            st.error("❌ OCR indisponível")
    with col3:
        st.success("✅ Excel Export")
    
    st.markdown("""
    ### 📋 O que será extraído:
    
    **Dados Gerais da Apólice:**
    - Informações do segurado (nome, CNPJ, endereço)
    - Dados da apólice (número, vigência, prêmios)
    - Informações do corretor
    
    **Para cada veículo (até 23 itens):**
    - ✅ **Identificação:** CEP, fabricante, modelo, ano, chassi, placa, combustível
    - ✅ **Características:** lotação, blindagem, kit gás, tipo de carroceria
    - ✅ **Coberturas:** limites e prêmios de todas as coberturas
    - ✅ **Franquias:** todos os tipos de vidros e componentes (66 campos por veículo)
    - ✅ **Classificação:** classe de bônus, código de identificação, FIPE
    
    **Total de colunas por veículo: ~66 campos**
    """)
    
    # Upload do arquivo
    uploaded_file = st.file_uploader(
        "Escolha um arquivo PDF da apólice Tokio Marine",
        type=['pdf'],
        help="PDFs com texto são processados instantaneamente. PDFs escaneados requerem OCR."
    )
    
    if uploaded_file is not None:
        # Mostra informações do arquivo
        st.success(f"✅ Arquivo carregado: {uploaded_file.name}")
        st.info(f"📊 Tamanho: {len(uploaded_file.getvalue())/1024:.1f} KB")
        
        # Botão para processar
        if st.button("🔄 Processar PDF Completo", type="primary"):
            # Extrai texto do PDF
            text = extract_text_from_pdf(uploaded_file)
            
            if text.strip():
                # Parse dos dados gerais
                with st.spinner("🧠 Analisando dados gerais da apólice..."):
                    dados_header = parse_header_data(text)
                
                # Parse dos veículos
                with st.spinner("🚗 Extraindo dados de todos os veículos..."):
                    vehicles = extract_vehicle_sections(text)
                    
                    all_vehicles_data = []
                    
                    # Barra de progresso para veículos
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    for i, vehicle in enumerate(vehicles):
                        progress = (i + 1) / len(vehicles)
                        progress_bar.progress(progress)
                        status_text.text(f"Processando veículo {vehicle['item']} de {len(vehicles)}...")
                        
                        vehicle_data = parse_vehicle_data(vehicle['content'], vehicle['item'])
                        all_vehicles_data.append(vehicle_data)
                    
                    progress_bar.progress(1.0)
                    status_text.text(f"✅ {len(vehicles)} veículos processados!")
                    
                    import time
                    time.sleep(1)
                    progress_bar.empty()
                    status_text.empty()
                
                # Mostra os dados extraídos
                st.markdown("## 📋 Dados Extraídos")
                
                # Estatísticas
                total_campos_header = len([v for v in dados_header.values() if v and v != ""])
                total_veiculos = len(all_vehicles_data)
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("📄 Campos Gerais", f"{total_campos_header}/20")
                with col2:
                    st.metric("🚗 Veículos Encontrados", total_veiculos)
                with col3:
                    if all_vehicles_data:
                        campos_por_veiculo = len([v for v in all_vehicles_data[0].values() if v and v != ""])
                        st.metric("📊 Campos por Veículo", f"{campos_por_veiculo}/66")
                
                # Dados gerais
                st.markdown("### 🏢 Informações Gerais da Apólice")
                dados_importantes = {
                    'Razão Social': dados_header.get('Razão Social', ''),
                    'CNPJ': dados_header.get('CNPJ', ''),
                    'Apólice': dados_header.get('Apólice', ''),
                    'Vigência do Seguro': dados_header.get('Vigência do Seguro', ''),
                    'Prêmio Total Geral': dados_header.get('Prêmio Total Geral', ''),
                    'Quantidade de Itens': dados_header.get('Quantidade de Itens', '')
                }
                
                for key, value in dados_importantes.items():
                    if value and value != "":
                        st.success(f"**{key}:** {value}")
                    else:
                        st.warning(f"**{key}:** Não encontrado")
                
                # Resumo dos veículos
                if all_vehicles_data:
                    st.markdown("### 🚙 Resumo dos Veículos")
                    
                    # Tabela resumo
                    resumo_veiculos = []
                    for vehicle in all_vehicles_data:
                        resumo = {
                            'Item': vehicle.get('Item', ''),
                            'Fabricante': vehicle.get('Fabricante', ''),
                            'Veículo': vehicle.get('Veículo', ''),
                            'Ano': vehicle.get('Ano Modelo', ''),
                            'Placa': vehicle.get('Placa', ''),
                            'Combustível': vehicle.get('Combustível', ''),
                            'Prêmio Total': vehicle.get('Prêmio Líquido Total', '')
                        }
                        resumo_veiculos.append(resumo)
                    
                    df_resumo = pd.DataFrame(resumo_veiculos)
                    st.dataframe(df_resumo, use_container_width=True)
                
                # Tabelas expandidas
                with st.expander("📊 Ver todos os dados detalhados"):
                    st.markdown("#### Dados Gerais Completos")
                    df_header_display = pd.DataFrame([dados_header])
                    st.dataframe(df_header_display, use_container_width=True)
                    
                    if all_vehicles_data:
                        st.markdown("#### Todos os Dados dos Veículos")
                        df_vehicles_display = pd.DataFrame(all_vehicles_data)
                        st.dataframe(df_vehicles_display, use_container_width=True)
                        
                        # Estatísticas por fabricante
                        if 'Fabricante' in df_vehicles_display.columns:
                            st.markdown("#### 📊 Estatísticas por Fabricante")
                            fabricantes = df_vehicles_display['Fabricante'].value_counts()
                            st.bar_chart(fabricantes)
                
                # Gera o arquivo Excel
                with st.spinner("📊 Gerando arquivo Excel..."):
                    excel_buffer = create_excel_file(dados_header, all_vehicles_data)
                
                # Botão de download
                st.markdown("## 💾 Download")
                apolice_numero = dados_header.get('Apólice', 'sem_numero')
                if not apolice_numero or apolice_numero == "":
                    apolice_numero = "sem_numero"
                nome_arquivo = f"apolice_completa_{apolice_numero}.xlsx"
                
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        label="📥 Baixar Planilha Excel Completa",
                        data=excel_buffer,
                        file_name=nome_arquivo,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                
                with col2:
                    st.info(f"""
                    **📋 Arquivo Excel contém:**
                    - Aba "Dados Gerais" (informações da apólice)
                    - Aba "Todos os Veículos" (dados completos)
                    - Aba "Dados Financeiros" (coberturas e prêmios)
                    - Aba "Franquias Vidros" (todas as franquias)
                    """)
                
                # Sucesso final
                if total_veiculos > 0:
                    taxa_sucesso = (total_campos_header/20 + campos_por_veiculo/66) / 2 * 100
                    if taxa_sucesso >= 80:
                        st.balloons()
                        st.success(f"🎉 Processamento concluído com excelência! {total_veiculos} veículos processados.")
                    else:
                        st.success(f"✅ Processamento concluído! {total_veiculos} veículos processados.")
                else:
                    st.warning("⚠️ Nenhum veículo foi encontrado no PDF. Verifique se é uma apólice Tokio Marine Auto Frota.")
                
            else:
                st.error("❌ Não foi possível extrair texto do PDF.")
        
        # Debug - Mostra preview do texto extraído
        with st.expander("🔍 Ver texto extraído do PDF (debug)"):
            if st.button("Extrair texto para visualização"):
                with st.spinner("Extraindo texto..."):
                    text_preview = extract_text_from_pdf(uploaded_file)
                if text_preview:
                    # Mostra seções de veículos encontradas
                    vehicles_debug = extract_vehicle_sections(text_preview)
                    st.info(f"🚗 Seções de veículos encontradas: {len(vehicles_debug)}")
                    
                    for i, vehicle in enumerate(vehicles_debug[:3]):  # Mostra apenas os primeiros 3
                        st.write(f"**Veículo {vehicle['item']}:**")
                        st.text_area(
                            f"Conteúdo do veículo {vehicle['item']}:", 
                            vehicle['content'][:500] + "..." if len(vehicle['content']) > 500 else vehicle['content'], 
                            height=200,
                            key=f"vehicle_{i}"
                        )
                    
                    if len(vehicles_debug) > 3:
                        st.info(f"... e mais {len(vehicles_debug) - 3} veículos")
                    
                    st.text_area(
                        "Texto completo extraído:", 
                        text_preview[:3000] + "..." if len(text_preview) > 3000 else text_preview, 
                        height=400
                    )
                    st.info(f"📏 Total de caracteres: {len(text_preview)}")
                else:
                    st.error("Não foi possível extrair texto do PDF")

    # Sidebar com informações
    with st.sidebar:
        st.markdown("## 🛠️ Configuração para Streamlit Cloud")
        
        st.markdown("**Para OCR funcionar, crie:**")
        
        st.markdown("📄 **packages.txt:**")
        st.code("""tesseract-ocr
tesseract-ocr-por
poppler-utils""", language="text")
        
        st.markdown("📄 **requirements.txt:**")
        st.code("""streamlit
PyPDF2
pandas
openpyxl
pytesseract
pdf2image
Pillow
numpy""", language="text")
        
        st.markdown("## 📊 Campos Extraídos")
        st.markdown("""
        **Por veículo (~66 campos):**
        - 🆔 Identificação (27 campos)
        - 💰 Coberturas (25 campos)  
        - 🔧 Franquias (17 campos)
        
        **Total:** Dados gerais + 23 veículos
        """)
        
        st.markdown("## 💡 Dicas")
        st.markdown("""
        - ⚡ PDFs com texto: Instantâneo
        - 📸 PDFs escaneados: Requer OCR
        - 🕐 Primeira vez com EasyOCR: Lento
        - 🚀 Use Tesseract quando possível
        - 📊 Resultado: 4 abas no Excel
        """)

if __name__ == "__main__":
    main()