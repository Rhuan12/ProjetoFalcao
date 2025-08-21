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
    """Extrai todas as seções de veículos do texto - versão robusta"""
    vehicles = []
    
    # Debug: vamos ver o que temos no texto
    st.info("🔍 Procurando seções de veículos no texto...")
    
    # Estratégia 1: Padrões mais flexíveis para "Descrição do Item"
    patterns_descricao = [
        r'Descrição do Item\s*-\s*(\d+)\s*-\s*Produto Auto Frota(.*?)(?=Descrição do Item\s*-\s*\d+\s*-\s*Produto Auto Frota|Assistência 24 Horas|\Z)',
        r'Descrição do Item\s*-\s*-\s*Produto Auto Frota(.*?)(?=Descrição do Item\s*-\s*[\d\s]*-\s*Produto Auto Frota|Assistência 24 Horas|\Z)',
        r'Item\s*-\s*(\d+)\s*-\s*Produto Auto Frota(.*?)(?=Item\s*-\s*\d+\s*-\s*Produto Auto Frota|Assistência 24 Horas|\Z)'
    ]
    
    for i, pattern in enumerate(patterns_descricao):
        sections = re.findall(pattern, text, re.DOTALL | re.IGNORECASE)
        
        if sections:
            st.success(f"✅ Encontradas {len(sections)} seções com padrão {i+1}")
            for j, section in enumerate(sections):
                if len(section) == 2 and section[0]:  # Pattern com número
                    item_num, content = section
                    vehicles.append({
                        'item': item_num.strip(),
                        'content': content.strip()
                    })
                elif len(section) == 2:  # Pattern sem número válido
                    content = section[1] if section[1] else section[0]
                    vehicles.append({
                        'item': str(j + 1),
                        'content': content.strip()
                    })
                else:  # Apenas conteúdo
                    content = section if isinstance(section, str) else section[0]
                    vehicles.append({
                        'item': str(j + 1),
                        'content': content.strip()
                    })
            break
    
    # Estratégia 2: Se não encontrou, procura por CEPs (indicativos de novos veículos)
    if not vehicles:
        st.info("🔍 Tentativa 2: Procurando por CEPs de pernoite...")
        cep_pattern = r'CEP de Pernoite do Veículo[:\s]*(\d{5}-?\d{3})(.*?)(?=CEP de Pernoite do Veículo|\Z)'
        cep_sections = re.findall(cep_pattern, text, re.DOTALL | re.IGNORECASE)
        
        if cep_sections:
            st.success(f"✅ Encontradas {len(cep_sections)} seções por CEP")
            for i, (cep, content) in enumerate(cep_sections):
                vehicles.append({
                    'item': str(i + 1),
                    'content': f"CEP de Pernoite do Veículo: {cep} {content}".strip()
                })
    
    # Estratégia 3: Procura por fabricantes como delimitadores
    if not vehicles:
        st.info("🔍 Tentativa 3: Procurando por fabricantes...")
        fabricantes = ['CHEVROLET', 'FORD', 'VOLKSWAGEN', 'FIAT', 'NISSAN', 'TOYOTA', 'BYD', 'MITSUBISHI']
        fabricante_pattern = r'Fabricante[:\s]*(' + '|'.join(fabricantes) + r')(.*?)(?=Fabricante[:\s]*(?:' + '|'.join(fabricantes) + r')|\Z)'
        
        fab_sections = re.findall(fabricante_pattern, text, re.DOTALL | re.IGNORECASE)
        
        if fab_sections:
            st.success(f"✅ Encontradas {len(fab_sections)} seções por fabricante")
            for i, (fabricante, content) in enumerate(fab_sections):
                vehicles.append({
                    'item': str(i + 1),
                    'content': f"Fabricante: {fabricante} {content}".strip()
                })
    
    # Estratégia 4: Procura por placas (último recurso)
    if not vehicles:
        st.info("🔍 Tentativa 4: Procurando por padrões de placas...")
        placa_pattern = r'Placa[:\s]*([A-Z]{3}\d{4}|[A-Z]{3}\d[A-Z]\d{2})(.*?)(?=Placa[:\s]*[A-Z]{3}[\dA-Z]|\Z)'
        
        placa_sections = re.findall(placa_pattern, text, re.DOTALL | re.IGNORECASE)
        
        if placa_sections:
            st.success(f"✅ Encontradas {len(placa_sections)} seções por placa")
            for i, (placa, content) in enumerate(placa_sections):
                vehicles.append({
                    'item': str(i + 1),
                    'content': f"Placa: {placa} {content}".strip()
                })
    
    if not vehicles:
        st.warning("⚠️ Nenhuma seção de veículo encontrada com os padrões conhecidos.")
        
        # Debug: mostra parte do texto para análise
        st.text_area("📝 Amostra do texto para debug:", text[:2000], height=200)
    
    return vehicles

def parse_vehicle_data(vehicle_content, item_num):
    """Extrai dados de um veículo específico"""
    
    def extract_money(patterns, text):
        """Extrai valores monetários"""
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
            if match:
                value = match.group(1).replace('.', '').replace(',', '.')
                try:
                    return float(value)
                except:
                    return value
        return ""
    
    def extract_simple(patterns, text):
        """Extrai valores simples"""
        result = extract_field(patterns, text)
        return result if result else ""
    
    dados = {
        # Identificação Básica
        'Item': item_num,
        'CEP de Pernoite': extract_simple([
            r'CEP de Pernoite do Veículo[:\s]*([^\s\n\r]+)',
            r'(\d{5}-?\d{3})'
        ], vehicle_content),
        
        'Fabricante': extract_simple([
            r'Fabricante[:\s]*([^\n\r]+?)(?=\s*Veículo[:\s]|$)',
            r'(CHEVROLET|FORD|VOLKSWAGEN|FIAT|NISSAN|TOYOTA|BYD|MITSUBISHI)'
        ], vehicle_content),
        
        'Veículo': extract_simple([
            r'Veículo[:\s]*([^\n\r]+?)(?=\s*(?:Ano Modelo|4º Eixo)|$)'
        ], vehicle_content),
        
        'Ano Modelo': extract_simple([
            r'Ano Modelo[:\s]*(\d{4})'
        ], vehicle_content),
        
        'Chassi': extract_simple([
            r'Chassi[:\s]*([A-Z0-9]{17})',
            r'Chassi[:\s]*([A-Z0-9]+)'
        ], vehicle_content),
        
        'Chassi Remarcado': extract_simple([
            r'Chassi Remarcado[:\s]*([^\n\r]+?)(?=\s*Placa[:\s]|$)'
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
        
        # Prêmio Líquido Total - campo mais importante
        'Prêmio Líquido Total': extract_money([
            r'Prêmio Líquido Total[:\s]*([0-9.,]+)',
            r'Prêmio Líquido Total.*?(\d+[\d.,]*)',
        ], vehicle_content),
    }
    
    # Extração simplificada de coberturas (valores principais)
    cobertura_patterns = {
        'Limite Colisão VMR': r'Valor Referenciado \(VMR\)\s*([0-9.,]+)',
        'Prêmio Colisão': r'Colisão, Incêndio e Roubo.*?([0-9.,]+)(?:\s+[0-9.,]+)?',
        'Limite RCF Danos Materiais': r'RCF-V - Danos Materiais\s+([0-9.,]+)',
        'Prêmio RCF Danos Materiais': r'RCF-V - Danos Materiais\s+[0-9.,]+\s+([0-9.,]+)',
        'Limite APP Morte': r'APP - Morte por Passageiro\s+([0-9.,]+)',
        'Prêmio APP Morte': r'APP - Morte por Passageiro\s+[0-9.,]+\s+([0-9.,]+)',
    }
    
    for field, pattern in cobertura_patterns.items():
        dados[field] = extract_money([pattern], vehicle_content)
    
    # Franquias principais
    franquia_patterns = {
        'Franquia Parabrisa': r'Parabrisa[:\s]*R\$\s*([0-9.,]+)',
        'Franquia Lateral': r'Lateral[:\s]*R\$\s*([0-9.,]+)',
        'Franquia Farol': r'Farol[^:]*[:\s]*R\$\s*([0-9.,]+)',
        'Franquia Retrovisor': r'Retrovisor[^:]*[:\s]*R\$\s*([0-9.,]+)',
    }
    
    for field, pattern in franquia_patterns.items():
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
        
        'Apólice': extract_field([
            r'Apólice[:\s]*([^\n\r]+?)(?=\s*Negócio|$)'
        ], text),
        
        'Vigência do Seguro': extract_field([
            r'Vigência do Seguro[:\s]*([^\n\r]+?)(?=\s*Data|$)'
        ], text),
        
        'Prêmio Total Geral': extract_field([
            r'Prêmio Total[:\s]*R\$\s*([^\n\r]+?)(?=\s*Cobrança|$)'
        ], text),
        
        'Quantidade de Itens': extract_field([
            r'Quantidade de Itens[:\s]*([^\n\r]+?)(?=\s*Sucursal|$)'
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
            
            # Aba resumo com campos principais
            resumo_columns = [
                'Item', 'Fabricante', 'Veículo', 'Ano Modelo', 'Placa', 'Chassi',
                'Combustível', 'Prêmio Líquido Total', 'Classe de Bônus'
            ]
            
            resumo_data = []
            for vehicle in all_vehicles_data:
                resumo_row = {col: vehicle.get(col, '') for col in resumo_columns}
                resumo_data.append(resumo_row)
            
            df_resumo = pd.DataFrame(resumo_data)
            df_resumo.to_excel(writer, sheet_name='Resumo Veículos', index=False)
    
    buffer.seek(0)
    return buffer

def main():
    st.title("🚗 Conversor de Apólices Tokio Marine")
    st.markdown("---")
    
    # Status do sistema
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
    ### 📋 Extração Robusta de Dados:
    
    **✅ Múltiplas estratégias de detecção:**
    - Padrão "Descrição do Item" (preferencial)
    - Detecção por CEP de pernoite
    - Detecção por fabricantes
    - Detecção por placas (fallback)
    
    **📊 Dados extraídos por veículo:**
    - Identificação completa (CEP, fabricante, modelo, etc.)
    - Valores de coberturas e prêmios
    - Franquias principais
    - Classe de bônus e códigos
    """)
    
    # Upload do arquivo
    uploaded_file = st.file_uploader(
        "Escolha um arquivo PDF da apólice Tokio Marine",
        type=['pdf'],
        help="O sistema tentará múltiplas estratégias para encontrar os veículos no PDF."
    )
    
    if uploaded_file is not None:
        st.success(f"✅ Arquivo carregado: {uploaded_file.name}")
        st.info(f"📊 Tamanho: {len(uploaded_file.getvalue())/1024:.1f} KB")
        
        if st.button("🔄 Processar PDF", type="primary"):
            # Extrai texto do PDF
            text = extract_text_from_pdf(uploaded_file)
            
            if text.strip():
                # Parse dos dados gerais
                with st.spinner("🏢 Analisando dados gerais da apólice..."):
                    dados_header = parse_header_data(text)
                
                # Parse dos veículos com estratégias múltiplas
                st.markdown("### 🔍 Buscando Veículos na Apólice")
                vehicles = extract_vehicle_sections(text)
                
                if vehicles:
                    with st.spinner("🚗 Extraindo dados dos veículos..."):
                        all_vehicles_data = []
                        
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
                
                # Mostra resultados
                st.markdown("## 📋 Dados Extraídos")
                
                # Métricas
                col1, col2, col3 = st.columns(3)
                with col1:
                    campos_header = len([v for v in dados_header.values() if v and v != ""])
                    st.metric("📄 Dados Gerais", f"{campos_header}/8")
                with col2:
                    st.metric("🚗 Veículos", len(vehicles))
                with col3:
                    if vehicles and all_vehicles_data:
                        campos_veiculo = len([v for v in all_vehicles_data[0].values() if v and v != ""])
                        st.metric("📊 Campos/Veículo", f"{campos_veiculo}/35")
                
                # Dados gerais
                st.markdown("### 🏢 Informações da Apólice")
                for key, value in dados_header.items():
                    if value and value != "":
                        st.success(f"**{key}:** {value}")
                    else:
                        st.warning(f"**{key}:** Não encontrado")
                
                # Resumo dos veículos
                if vehicles and all_vehicles_data:
                    st.markdown("### 🚙 Veículos Encontrados")
                    
                    resumo_data = []
                    for vehicle in all_vehicles_data:
                        resumo = {
                            'Item': vehicle.get('Item', ''),
                            'Fabricante': vehicle.get('Fabricante', ''),
                            'Veículo': vehicle.get('Veículo', ''),
                            'Ano': vehicle.get('Ano Modelo', ''),
                            'Placa': vehicle.get('Placa', ''),
                            'Prêmio Total': vehicle.get('Prêmio Líquido Total', '')
                        }
                        resumo_data.append(resumo)
                    
                    df_resumo = pd.DataFrame(resumo_data)
                    st.dataframe(df_resumo, use_container_width=True)
                    
                    # Estatísticas por fabricante
                    if 'Fabricante' in df_resumo.columns:
                        fabricantes_count = df_resumo['Fabricante'].value_counts()
                        if len(fabricantes_count) > 0:
                            st.markdown("#### 📊 Distribuição por Fabricante")
                            st.bar_chart(fabricantes_count)
                
                # Tabelas expandidas
                with st.expander("📊 Ver todos os dados detalhados"):
                    st.markdown("#### Dados Gerais Completos")
                    df_header_display = pd.DataFrame([dados_header])
                    st.dataframe(df_header_display, use_container_width=True)
                    
                    if vehicles and all_vehicles_data:
                        st.markdown("#### Todos os Dados dos Veículos")
                        df_vehicles_display = pd.DataFrame(all_vehicles_data)
                        st.dataframe(df_vehicles_display, use_container_width=True)
                        
                        # Análise de completude dos dados
                        st.markdown("#### 🎯 Análise de Completude dos Dados")
                        
                        # Calcular estatísticas de completude
                        total_campos = len(df_vehicles_display.columns)
                        completude_por_veiculo = []
                        
                        for index, row in df_vehicles_display.iterrows():
                            campos_preenchidos = sum(1 for val in row if val and str(val).strip() and str(val) != '')
                            percentual = (campos_preenchidos / total_campos) * 100
                            completude_por_veiculo.append({
                                'Veículo': f"Item {row.get('Item', index+1)}",
                                'Campos Preenchidos': f"{campos_preenchidos}/{total_campos}",
                                'Percentual': f"{percentual:.1f}%"
                            })
                        
                        df_completude = pd.DataFrame(completude_por_veiculo)
                        st.dataframe(df_completude, use_container_width=True)
                
                # Gera o arquivo Excel
                if vehicles and all_vehicles_data:
                    with st.spinner("📊 Gerando arquivo Excel..."):
                        excel_buffer = create_excel_file(dados_header, all_vehicles_data)
                    
                    # Seção de download
                    st.markdown("## 💾 Download")
                    
                    apolice_numero = dados_header.get('Apólice', '')
                    if not apolice_numero or apolice_numero.strip() == "":
                        apolice_numero = "sem_numero"
                    nome_arquivo = f"tokio_marine_apolice_{apolice_numero}.xlsx"
                    
                    col1, col2 = st.columns([2, 1])
                    
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
                        - 📄 Dados Gerais da Apólice
                        - 🚗 Todos os Veículos ({len(vehicles)} itens)
                        - 📊 Resumo dos Veículos
                        """)
                    
                    # Sucesso final
                    taxa_deteccao = len(vehicles)
                    if taxa_deteccao > 20:
                        st.balloons()
                        st.success(f"🎉 Excelente! {taxa_deteccao} veículos detectados e processados com sucesso!")
                    elif taxa_deteccao > 10:
                        st.success(f"✅ Muito bom! {taxa_deteccao} veículos detectados e processados!")
                    elif taxa_deteccao > 0:
                        st.success(f"✅ {taxa_deteccao} veículos detectados e processados!")
                    else:
                        st.warning("⚠️ Nenhum veículo detectado. Verifique se é uma apólice Tokio Marine Auto Frota.")
                        
                else:
                    st.warning("⚠️ Nenhum veículo foi encontrado no PDF.")
                    st.info("💡 Isso pode acontecer se:")
                    st.markdown("""
                    - O PDF não é uma apólice Tokio Marine Auto Frota
                    - O formato do PDF é muito diferente do esperado
                    - A qualidade do OCR (se usado) foi insuficiente
                    - O arquivo está corrompido ou protegido
                    """)
                
            else:
                st.error("❌ Não foi possível extrair texto do PDF.")
                st.info("💡 Possíveis causas:")
                st.markdown("""
                - PDF protegido ou criptografado
                - PDF corrompido
                - OCR não disponível para PDFs de imagem
                - Formato de arquivo não suportado
                """)
        
        # Debug - Seção expandível para análise do texto
        with st.expander("🔍 Debug: Ver texto extraído e análise detalhada"):
            if st.button("🔍 Extrair e Analisar Texto"):
                with st.spinner("Extraindo texto para análise..."):
                    debug_text = extract_text_from_pdf(uploaded_file)
                
                if debug_text:
                    # Estatísticas do texto
                    st.markdown("#### 📊 Estatísticas do Texto Extraído")
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.metric("Caracteres", f"{len(debug_text):,}")
                    with col2:
                        palavras = len(debug_text.split())
                        st.metric("Palavras", f"{palavras:,}")
                    with col3:
                        linhas = len(debug_text.split('\n'))
                        st.metric("Linhas", f"{linhas:,}")
                    with col4:
                        # Conta menções de "Descrição do Item"
                        mencoes_item = len(re.findall(r'Descrição do Item', debug_text, re.IGNORECASE))
                        st.metric("'Descrição do Item'", mencoes_item)
                    
                    # Busca por padrões específicos
                    st.markdown("#### 🔍 Padrões Encontrados")
                    
                    patterns_debug = {
                        "CEPs encontrados": r'\d{5}-?\d{3}',
                        "Placas encontradas": r'[A-Z]{3}\d{4}|[A-Z]{3}\d[A-Z]\d{2}',
                        "Fabricantes encontrados": r'(CHEVROLET|FORD|VOLKSWAGEN|FIAT|NISSAN|TOYOTA|BYD|MITSUBISHI)',
                        "Valores R$ encontrados": r'R\$\s*[\d.,]+',
                        "Anos encontrados": r'\b(19|20)\d{2}\b'
                    }
                    
                    for desc, pattern in patterns_debug.items():
                        matches = re.findall(pattern, debug_text, re.IGNORECASE)
                        if matches:
                            st.success(f"**{desc}:** {len(matches)} - Exemplos: {', '.join(matches[:5])}")
                        else:
                            st.warning(f"**{desc}:** Nenhum encontrado")
                    
                    # Tenta detectar veículos com debug
                    st.markdown("#### 🚗 Debug: Detecção de Veículos")
                    debug_vehicles = extract_vehicle_sections(debug_text)
                    
                    if debug_vehicles:
                        st.success(f"✅ {len(debug_vehicles)} seções de veículos detectadas!")
                        
                        # Mostra primeiras seções encontradas
                        for i, vehicle in enumerate(debug_vehicles[:3]):
                            st.markdown(f"**🚗 Veículo {vehicle['item']}:**")
                            content_preview = vehicle['content'][:800] + "..." if len(vehicle['content']) > 800 else vehicle['content']
                            st.text_area(
                                f"Conteúdo do Item {vehicle['item']}:", 
                                content_preview, 
                                height=200,
                                key=f"debug_vehicle_{i}"
                            )
                        
                        if len(debug_vehicles) > 3:
                            st.info(f"... e mais {len(debug_vehicles) - 3} veículos detectados.")
                    
                    # Amostra do texto completo
                    st.markdown("#### 📝 Amostra do Texto Completo")
                    texto_amostra = debug_text[:5000] + "\n\n... (texto truncado)" if len(debug_text) > 5000 else debug_text
                    st.text_area(
                        "Texto extraído completo:", 
                        texto_amostra, 
                        height=300
                    )
                    
                else:
                    st.error("❌ Não foi possível extrair texto para debug")

    # Sidebar com informações técnicas
    with st.sidebar:
        st.markdown("## 🛠️ Informações Técnicas")
        
        st.markdown("### 📋 Estratégias de Detecção")
        st.markdown("""
        **1. Padrão Principal:** "Descrição do Item"
        **2. Fallback CEP:** CEP de pernoite  
        **3. Fallback Fabricante:** Marcas conhecidas
        **4. Fallback Placa:** Padrões de placas
        """)
        
        st.markdown("### 🔧 Configuração OCR")
        st.markdown("**Para Streamlit Cloud:**")
        
        st.code("""# packages.txt
tesseract-ocr
tesseract-ocr-por
poppler-utils""", language="text")
        
        st.code("""# requirements.txt
streamlit
PyPDF2
pandas
openpyxl
pytesseract
pdf2image
Pillow
numpy""", language="text")
        
        st.markdown("### 📊 Campos Extraídos")
        st.markdown("""
        **Por veículo:**
        - 🆔 Identificação (15+ campos)
        - 💰 Coberturas (10+ campos)  
        - 🔧 Franquias (10+ campos)
        
        **Total:** ~35 campos/veículo
        """)
        
        st.markdown("### 💡 Dicas de Uso")
        st.markdown("""
        - ⚡ PDFs nativos: Processamento instantâneo
        - 📸 PDFs escaneados: Requer OCR (mais lento)
        - 🔍 Use o debug se não encontrar veículos
        - 📊 Resultado em múltiplas abas Excel
        - 🎯 Sistema detecta automaticamente o formato
        """)
        
        st.markdown("### 📞 Suporte")
        st.markdown("""
        Se não conseguir extrair os dados:
        1. Verifique se é uma apólice Tokio Marine
        2. Use a função de debug
        3. Teste com PDF de melhor qualidade
        4. Verifique se o OCR está funcionando
        """)

if __name__ == "__main__":
    main()