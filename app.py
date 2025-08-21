import streamlit as st
import PyPDF2
import pandas as pd
import re
import os
from io import BytesIO
import tempfile

# Configuração da página
st.set_page_config(
    page_title="Conversor de Apólices Tokio Marine",
    page_icon="📄",
    layout="wide"
)

def extract_text_from_pdf(pdf_file):
    """
    Extrai texto de um arquivo PDF usando PyPDF2
    """
    try:
        # Cria um arquivo temporário
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(pdf_file.read())
            tmp_file_path = tmp_file.name
        
        text = ""
        with open(tmp_file_path, "rb") as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                text += page.extract_text() + "\n"
        
        # Remove o arquivo temporário
        os.unlink(tmp_file_path)
        
        return text
    except Exception as e:
        st.error(f"Erro ao extrair texto do PDF: {e}")
        return ""

def extract_field(patterns, text):
    """
    Procura por uma lista de padrões regex e retorna o valor encontrado
    """
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return "Não encontrado"

def parse_tokio_data(text):
    """
    Extrai dados específicos da apólice Tokio Marine
    """
    # Dados do cabeçalho/cliente
    dados_header = {
        "NOME DO CLIENTE": extract_field([r"Proprietário[:\s]*(.*)"], text),
        "CNPJ": extract_field([r"CNPJ[:\s]*(.*)"], text),
        "APÓLICE": extract_field([r"Nr Apólice.*?(\d+)"], text),
        "VIGÊNCIA": extract_field([r"Venc Apólice.*?:\s*([\d/]+)"], text),
    }

    # Dados do veículo
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
        "TIPO DE SEGURO": "Renovação Tokio sem sinistro",
        "NR APÓLICE CONGENERE": extract_field([r"Nr Apólice Congenere:\s*(.*)"], text),
        "NOME DA CONGENERE": extract_field([r"Nome da Congenere:\s*(.*)"], text),
        "VENC APÓLICE CONGENERE": extract_field([r"Venc Apólice Cong.: (.*)"], text),
        "CLASSE DE BÔNUS": extract_field([r"Classe de Bônus:\s*(.*)"], text),
        "CÓDIGO DE IDENTIFICAÇÃO (CI)": extract_field([r"Código de Identificação \(CI\):\s*(.*)"], text),
        "KM DE REBOQUE": extract_field([r"Km de Reboque:\s*(.*)"], text),
        "KM (ADICIONAL)": extract_field([r"km\(Adicional\):\s*(.*)"], text),
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
    1. Faça o upload da sua apólice em PDF
    2. Aguarde o processamento automático
    3. Visualize os dados extraídos
    4. Baixe a planilha Excel gerada
    """)
    
    # Upload do arquivo
    uploaded_file = st.file_uploader(
        "Escolha um arquivo PDF da apólice Tokio Marine",
        type=['pdf'],
        help="Faça upload do arquivo PDF da apólice para conversão"
    )
    
    if uploaded_file is not None:
        # Mostra informações do arquivo
        st.success(f"✅ Arquivo carregado: {uploaded_file.name}")
        st.info(f"📊 Tamanho: {len(uploaded_file.getvalue())/1024:.1f} KB")
        
        # Botão para processar
        if st.button("🔄 Processar PDF", type="primary"):
            with st.spinner("Extraindo dados da apólice..."):
                # Extrai texto do PDF
                text = extract_text_from_pdf(uploaded_file)
                
                if text.strip():
                    # Parse dos dados
                    dados_header, dados_veiculo = parse_tokio_data(text)
                    
                    # Mostra os dados extraídos
                    st.markdown("## 📋 Dados Extraídos")
                    
                    # Dados gerais em duas colunas
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.markdown("### 🏢 Informações do Cliente")
                        for key, value in dados_header.items():
                            st.text(f"{key}: {value}")
                    
                    with col2:
                        st.markdown("### 🚙 Informações do Veículo")
                        # Mostra apenas os primeiros campos para não sobrecarregar
                        campos_importantes = [
                            "VEÍCULO", "ANO MODELO", "PLACA", "CHASSI", 
                            "COMBUSTÍVEL", "FIPE", "CLASSE DE BÔNUS"
                        ]
                        for campo in campos_importantes:
                            if campo in dados_veiculo:
                                st.text(f"{campo}: {dados_veiculo[campo]}")
                    
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
                    nome_arquivo = f"apolice_{dados_header.get('APÓLICE', 'sem_numero')}.xlsx"
                    
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
                text_preview = extract_text_from_pdf(uploaded_file)
                if text_preview:
                    st.text_area("Texto extraído:", text_preview[:2000] + "..." if len(text_preview) > 2000 else text_preview, height=300)
                else:
                    st.error("Não foi possível extrair texto do PDF")

if __name__ == "__main__":
    main()