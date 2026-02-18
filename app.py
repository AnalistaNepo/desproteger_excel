import streamlit as st
import os
import zipfile
import shutil
import re
import tempfile
from io import BytesIO

# Configuração da página
st.set_page_config(
    page_title="Desproteger Excel",
    page_icon="🔓",
    layout="centered"
)

# Estilo CSS personalizado
st.markdown("""
<style>
    .reportview-container {
        background: #f0f2f6;
    }
    .main {
        background-color: #ffffff;
        padding: 2rem;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    h1 {
        color: #1f2937;
        font-family: 'Helvetica Neue', sans-serif;
        text-align: center;
        font-weight: 700;
        margin-bottom: 2rem;
    }
    .stButton>button {
        width: 100%;
        background-color: #4CAF50;
        color: white;
        border: none;
        padding: 0.5rem 1rem;
        border-radius: 5px;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #45a049;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2);
    }
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: #f8f9fa;
        color: #6c757d;
        text-align: center;
        padding: 10px;
        font-size: 0.8rem;
        border-top: 1px solid #e9ecef;
    }
    .uploaded-file {
        border: 2px dashed #ced4da;
        border-radius: 5px;
        padding: 1rem;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

def process_excel(uploaded_file):
    """
    Processa o arquivo Excel enviado para remover a proteção.
    Retorna os bytes do arquivo desprotegido.
    """
    # Criar diretório temporário para processamento
    with tempfile.TemporaryDirectory() as temp_dir:
        # Caminhos temporários
        temp_input_path = os.path.join(temp_dir, "input.xlsx")
        temp_extract_dir = os.path.join(temp_dir, "extracted")
        temp_output_path = os.path.join(temp_dir, "output.xlsx")
        
        # Salvar arquivo enviado
        with open(temp_input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Extrair arquivos
        with zipfile.ZipFile(temp_input_path, 'r') as zip_ref:
            zip_ref.extractall(temp_extract_dir)
            
        # Processar planilhas
        worksheets_dir = os.path.join(temp_extract_dir, "xl", "worksheets")
        if os.path.exists(worksheets_dir):
            for sheet_file in os.listdir(worksheets_dir):
                if sheet_file.endswith(".xml"):
                    full_path = os.path.join(worksheets_dir, sheet_file)
                    with open(full_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    # Remover proteção da planilha
                    new_content = re.sub(r'<sheetProtection[^>]*/>', '', content)
                    
                    if content != new_content:
                        with open(full_path, 'w', encoding='utf-8') as f:
                            f.write(new_content)

        # Processar workbook.xml
        workbook_path = os.path.join(temp_extract_dir, "xl", "workbook.xml")
        if os.path.exists(workbook_path):
            with open(workbook_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Remover proteção da pasta de trabalho
            new_content = re.sub(r'<workbookProtection[^>]*/>', '', content)
            new_content = re.sub(r'<workbookProtection[^>]*>.*?</workbookProtection>', '', new_content, flags=re.DOTALL)

            if content != new_content:
                with open(workbook_path, 'w', encoding='utf-8') as f:
                    f.write(new_content)

        # Compactar novamente
        with zipfile.ZipFile(temp_output_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for root, dirs, files in os.walk(temp_extract_dir):
                for file in files:
                    file_path_on_disk = os.path.join(root, file)
                    archive_name = os.path.relpath(file_path_on_disk, temp_extract_dir)
                    zip_out.write(file_path_on_disk, archive_name)
                    
        # Ler o arquivo gerado para retornar bytes
        with open(temp_output_path, "rb") as f:
            return f.read()

def main():
    st.title("🔓 Desproteger Excel")
    st.markdown("Remova senhas de proteção de planilhas e pastas de trabalho de forma simples e rápida.")

    uploaded_file = st.file_uploader("Arraste e solte seu arquivo Excel aqui", type=['xlsx', 'xlsm'])

    if uploaded_file is not None:
        st.success(f"Arquivo carregado: {uploaded_file.name}")
        
        if st.button("Desproteger Arquivo"):
            with st.spinner("Processando arquivo..."):
                try:
                    protected_bytes = process_excel(uploaded_file)
                    
                    file_name_output = f"{os.path.splitext(uploaded_file.name)[0]}_unprotected.xlsx"
                    
                    st.balloons()
                    st.success("Arquivo desprotegido com sucesso!")
                    
                    st.download_button(
                        label="📥 Baixar Arquivo Desprotegido",
                        data=protected_bytes,
                        file_name=file_name_output,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                except Exception as e:
                    st.error(f"Ocorreu um erro ao processar o arquivo: {str(e)}")

    st.markdown('<div class="footer">Desenvolvido por Leonir Martins</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
