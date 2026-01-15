import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import io

# ==========================================
# CONFIGURAÇÕES E CHAVES (PREENCHA AQUI!)
# ==========================================
st.set_page_config(page_title="Solar Force", page_icon="🔴", layout="centered")

# 1. Chave do site ImgBB
IMGBB_API_KEY = "775d60bb1bcd4c621f61f0213e10ad7c" 

# 2. Configurações de E-mail
EMAIL_REMETENTE = "natan.10.dez@gmail.com" 
SENHA_EMAIL = "boyk xzgz vand oaxf" 
EMAIL_DESTINATARIO = "natan.10.dez@hotmail.com" 

# Senha Admin
SENHA_ADMIN = "solar2024"

# ==========================================
# DESIGN E ESTILO
# ==========================================
st.markdown("""
    <style>
    /* Botão Vermelho Coca-Cola */
    div.stButton > button:first-child {
        background-color: #F40009 !important;
        color: white !important;
        border-radius: 12px;
        width: 100%;
        font-weight: bold;
    }
    /* Esconde o olho da senha */
    button[aria-label="Show password"] {
        display: none !important;
    }
    /* Fontes */
    .stTextInput label, .stMultiSelect label, .stTextArea label, .stFileUploader label {
        font-size: 16px;
        font-weight: 600;
    }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# FUNÇÕES DE BACK-END
# ==========================================

def get_google_sheet(nome_da_aba):
    """Conecta em uma aba específica da planilha"""
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
             "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("credenciais.json", scope)
    client = gspread.authorize(creds)
    
    # Abre a planilha principal e seleciona a aba pelo nome
    planilha = client.open("Sistema Solar Force - Dados")
    
    if nome_da_aba == "Principal":
        return planilha.sheet1
    else:
        # Certifique-se que criou a aba 'Controle_GDM' no Google Sheets
        return planilha.worksheet(nome_da_aba)

def upload_imagem(arquivo):
    try:
        url = "https://api.imgbb.com/1/upload"
        payload = {"key": IMGBB_API_KEY, "expiration": 0}
        files = {"image": arquivo.getvalue()}
        response = requests.post(url, data=payload, files=files)
        return response.json()["data"]["url"]
    except Exception as e:
        return f"[Erro: {e}]"

def salvar_no_google(dados, nome_aba):
    """Salva dados na aba especificada"""
    sheet = get_google_sheet(nome_aba)
    sheet.append_row(dados)

def enviar_relatorio_email(tipo_relatorio):
    """Gera Excel e envia email baseado no tipo (Geral ou GDM)"""
    try:
        # Define qual aba ler e qual assunto usar
        if tipo_relatorio == "Geral":
            sheet = get_google_sheet("Principal")
            assunto = "Resumo Consolidado - VISITAS"
            nome_arquivo = "Relatorio_Visitas"
        elif tipo_relatorio == "GDM":
            sheet = get_google_sheet("Controle_GDM")
            assunto = "Resumo Consolidado - CONTROLE GDM"
            nome_arquivo = "Relatorio_GDM"

        dados = sheet.get_all_records()
        df = pd.DataFrame(dados)
        
        if df.empty:
            return "Vazio"

        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Relatorio')
        buffer_excel.seek(0)

        msg = MIMEMultipart()
        msg['From'] = EMAIL_REMETENTE
        msg['To'] = EMAIL_DESTINATARIO
        msg['Subject'] = f"{assunto} - Solar Force ({datetime.now().strftime('%d/%m')})"

        body = f"""
        Olá,
        
        Segue em anexo o relatório solicitado: {assunto}.
        Total de registros: {len(df)}
        
        Atenciosamente,
        Sistema Solar Force
        """
        msg.attach(MIMEText(body, 'plain'))

        part = MIMEBase('application', 'octet-stream')
        part.set_payload(buffer_excel.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename={nome_arquivo}_{datetime.now().strftime('%d_%m')}.xlsx")
        msg.attach(part)

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_REMETENTE, SENHA_EMAIL)
        server.sendmail(EMAIL_REMETENTE, EMAIL_DESTINATARIO, msg.as_string())
        server.quit()
        return "Sucesso"

    except Exception as e:
        return f"Erro Técnico: {str(e)}"

# ==========================================
# INTERFACE (FRONT-END)
# ==========================================

# Menu com 3 Opções
menu = st.sidebar.selectbox("Navegação", [
    "Área do Promotor (Visitas)", 
    "Controle de GDM ❄️", 
    "Painel Administrativo"
])

st.image("https://upload.wikimedia.org/wikipedia/commons/c/ce/Coca-Cola_logo.svg", width=180)

# --- OPÇÃO 1: VISITAS (PROMOTOR) ---
if menu == "Área do Promotor (Visitas)":
    st.markdown("<h1 style='text-align: center;'>Relatório de Campo</h1>", unsafe_allow_html=True)
    st.info("Preencha os dados da visita diária.")

    with st.form(key="form_visita"):
        col1, col2 = st.columns(2)
        with col1:
            nome = st.text_input("Nome", placeholder="Ex: João Silva") 
        with col2:
            matricula = st.text_input("Matrícula", placeholder="Ex: 123456")

        col3, col4 = st.columns(2)
        with col3:
            cod_loja = st.text_input("Código Loja", placeholder="Ex: 9988")
        with col4:
            cidade = st.text_input("Cidade", placeholder="Ex: Belém")
        
        missoes = st.multiselect("Atividades", 
            ["Pesquisa Red", "Red Simulado", "Inventário GDM", "Troca GDM", "Manutenção", "Troca de EPI's", "Solicitação de Crachá", "Outros"])
        
        obs = st.text_area("Observações")
        arquivos_fotos = st.file_uploader("Evidências (Opcional)", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        submit = st.form_submit_button("REGISTRAR VISITA 💾")

    if submit:
        if not nome or not cod_loja or not missoes:
            st.error("⚠️ Preencha Nome, Loja e Atividades!")
        else:
            with st.spinner('Enviando...'):
                try:
                    # Upload Múltiplo
                    lista_links = []
                    if arquivos_fotos:
                        for arquivo in arquivos_fotos:
                            lista_links.append(upload_imagem(arquivo))
                        link_final = " | ".join(lista_links)
                    else:
                        link_final = "-"
                    
                    data_hora = datetime.now().strftime("%d/%m/%Y %H:%M")
                    missoes_txt = ", ".join(missoes)
                    
                    # Salva na aba Principal (Sheet1)
                    salvar_no_google([data_hora, nome, matricula, cod_loja, cidade, missoes_txt, obs, link_final], "Principal")
                    
                    st.success("✅ Visita registrada com sucesso!")
                except Exception as e:
                    st.error(f"Erro: {e}")

# --- OPÇÃO 2: CONTROLE DE GDM (CORRIGIDO!) ---
elif menu == "Controle de GDM ❄️":
    st.markdown("<h1 style='text-align: center;'>Controle de GDM</h1>", unsafe_allow_html=True)
    st.warning("Use esta área apenas para reportar divergências em Geladeiras.")

    with st.form(key="form_gdm"):
        # Identificação Básica
        col1, col2 = st.columns(2)
        with col1:
            nome = st.text_input("Nome Promotor") # Variável correta: nome
        with col2:
            cod_loja = st.text_input("Código Loja")
            
        st.markdown("### 🧊 Status das GDMs")
        st.caption("Insira os códigos patrimoniais separados por vírgula ou espaço.")
        
        gdm_nao_pesq = st.text_area("GDMs Não Pesquisadas (Estão no PDV mas não no sistema)", height=80)
        gdm_perdidas = st.text_area("GDMs Perdidas (Estão no sistema mas não no PDV)", height=80)
        gdm_paradas = st.text_area("GDMs Paradas/Quebradas", height=80)
        
        obs_gdm = st.text_input("Observação Geral")
        fotos_gdm = st.file_uploader("Foto da Etiqueta/GDM", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        submit_gdm = st.form_submit_button("REGISTRAR GDM ❄️")
        
    if submit_gdm:
        if not nome or not cod_loja:
            st.error("⚠️ Identifique o promotor e a loja!")
        elif not (gdm_nao_pesq or gdm_perdidas or gdm_paradas):
            st.error("⚠️ Preencha pelo menos um campo de GDM!")
        else:
            with st.spinner('Registrando GDM...'):
                try:
                    # Upload Fotos GDM
                    lista_links = []
                    if fotos_gdm:
                        for arquivo in fotos_gdm:
                            lista_links.append(upload_imagem(arquivo))
                        link_final_gdm = " | ".join(lista_links)
                    else:
                        link_final_gdm = "-"
                        
                    data_hora = datetime.now().strftime("%d/%m/%Y %H:%M")
                    
                    # CORREÇÃO APLICADA: Trocado 'nome_promotor' por 'nome'
                    salvar_no_google([data_hora, nome, cod_loja, gdm_nao_pesq, gdm_perdidas, gdm_paradas, obs_gdm, link_final_gdm], "Controle_GDM")
                    
                    st.success("✅ Ocorrência de GDM registrada!")
                except Exception as e:
                    st.error(f"Erro ao salvar: {e}. Verifique se a aba 'Controle_GDM' existe na planilha.")

# --- OPÇÃO 3: ADMINISTRAÇÃO ---
elif menu == "Painel Administrativo":
    st.markdown("<h1 style='text-align: center;'>Painel Gerencial</h1>", unsafe_allow_html=True)
    st.markdown("---")
    
    senha_input = st.text_input("🔑 Senha de administrador:", type="password")
    
    if senha_input == SENHA_ADMIN:
        st.success("Painel Liberado")
        
        col_A, col_B = st.columns(2)
        
        # --- BOTÃO 1: RELATÓRIO GERAL DE VISITAS ---
        with col_A:
            st.info("📋 **Relatório de Visitas**")
            st.caption("Puxa dados da aba Principal.")
            if st.button("Enviar Relatório VISITAS 📧"):
                with st.spinner("Processando Visitas..."):
                    res = enviar_relatorio_email("Geral")
                    if res == "Sucesso": st.success("Enviado!")
                    elif res == "Vazio": st.warning("Sem dados.")
                    else: st.error(res)

        # --- BOTÃO 2: RELATÓRIO DE GDM ---
        with col_B:
            st.info("❄️ **Relatório de GDM**")
            st.caption("Puxa dados da aba Controle_GDM.")
            if st.button("Enviar Relatório GDM 📧"):
                with st.spinner("Processando GDMs..."):
                    res = enviar_relatorio_email("GDM")
                    if res == "Sucesso": st.success("Enviado!")
                    elif res == "Vazio": st.warning("Sem dados.")
                    else: st.error(res)
    
    elif senha_input:
        st.error("Senha Incorreta.")