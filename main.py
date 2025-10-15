import streamlit as st
from streamlit_option_menu import option_menu  # Biblioteca externa
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from googleapiclient.discovery import build
import tempfile
from googleapiclient.http import MediaFileUpload
from fpdf import FPDF
from fpdf.enums import Align
import os
from datetime import date
import json

#biblioecas para mandar mensagens
#import smtplib
#import ssl
#import pywhatkit as kit
#from selenium import webdriver
#from selenium.webdriver.common.by import By
#from selenium.webdriver.common.keys import Keys
#import urllib
#import time


#import plotly as px

st.set_page_config(
    page_title="ATLETA PRO",
    page_icon="🏃",
    layout="wide",
    initial_sidebar_state="collapsed"
)

## Configurações para acesso ao sheets e drive
filename = "internos\\atletapro-470813-54d37a8d09fe.json" #file para acessar google sheets dentro da pasta internos
scopes = [
    "https://spreadsheets.google.com/feeds", 
    "https://www.googleapis.com/auth/drive",
]

if 'gcp_service_account_json' not in st.secrets:
    st.error("JSON da service account não encontrado nos secrets!")
    st.stop()

else:
    service_account_json = st.secrets["gcp_service_account_json"]
    service_account_info = json.loads(service_account_json)
    
    creds = ServiceAccountCredentials.from_json_keyfile_dict(
        keyfile_dict=service_account_info,
        scopes=scopes
    )

drive_service = build('drive', 'v3', credentials=creds)

def salvar_pdf_no_drive(pdf, nome_arquivo, pasta_id):
    """Salva PDF em Shared Drive (solução recomendada)"""
    try:
        # Salvar temporariamente
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
            pdf.output(temp_file.name)
            temp_path = temp_file.name
        
        # Metadados
        file_metadata = {
            'name': nome_arquivo,
            'parents': [pasta_id]
        }
        
        # Upload com suporte a Shared Drives
        media = MediaFileUpload(temp_path, mimetype='application/pdf', resumable=True)
        
        file = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            supportsAllDrives=True,  # ← CRÍTICO
            fields='id, name, webViewLink, webContentLink'
        ).execute()
        
        # Limpeza
        os.unlink(temp_path)
        
        st.success(f"✅ PDF ESTOQUE salva com sucesso!")
        st.write(f"**Arquivo:** {file['name']}")
        
        if file.get('webViewLink'):
            st.markdown(f"**🔗 [Abrir no Drive]({file['webViewLink']})**")
        
        return file
        
    except Exception as e:
        st.error(f"❌ Erro ao salvar: {e}")
        # Debug adicional
        st.write("⚠️ Certifique-se que:")
        st.write("- O Shared Drive existe")
        st.write("- A Service Account tem acesso como 'Editor'")
        st.write("- Você está usando o ID correto do Shared Drive")
        return None

client = gspread.authorize(creds)

planilha_completa = client.open(
    title="Gestão de Estoque - Atleta Pro", 
    folder_id="1j7PyEmL4CiTJReuIfyul7_DlSgwW5K0U"
    )

planilhaEstoque = planilha_completa.get_worksheet(0)

dados_estoque = planilhaEstoque.get_all_records() 

dados = pd.DataFrame(dados_estoque)

#Função de Status do estoque ("Atenção" caso abaixo)
def status():
    for produtos in dados["Código"]:
        linha_status = dados[dados['Código'] == produtos]
        

        quantidadeStatus = linha_status["Quantidade Atual"].unique().tolist()
        esStatus = linha_status["Estoque de Segurança"].unique().tolist()
        ordemdeCompraStatus = linha_status["PO"].unique().tolist()
        statusStatus = linha_status["Status"].unique().tolist()

        # Verificar se as listas não estão vazias
        if not all([quantidadeStatus, esStatus, ordemdeCompraStatus, statusStatus]):
            continue
            
        # Primeira condição: Se não for semiacabado e PO vazia
        #st.write(esStatus[0] != " - ", (ordemdeCompraStatus[0] == "" or ordemdeCompraStatus[0] == " - "))
        if esStatus[0] != " - " and (ordemdeCompraStatus[0] == "" or ordemdeCompraStatus[0] == " - "):
            try:
                
                if quantidadeStatus[0] < esStatus[0]:
                    if statusStatus[0] != "Atenção":
                        indiceStatus = dados.index[dados['Código'] == produtos].tolist()
                        linhaStatus = int(indiceStatus[0]) + 2
                        planilhaEstoque.update_cell(row=int(linhaStatus), col=14, value="Atenção")
    
                        if statusStatus[0] != " - ":
                            planilhaEstoque.update_cell(row=int(linhaStatus), col=15, value=" - ")
                    
                else:
                    if statusStatus[0] != " - ":
                        indiceStatus = dados.index[dados['Código'] == produtos].tolist()
                        linhaStatus = int(indiceStatus[0]) + 2
                        planilhaEstoque.update_cell(row=int(linhaStatus), col=14, value=" - ")

                        if statusStatus[0] != " - ":
                            planilhaEstoque.update_cell(row=int(linhaStatus), col=15, value=" - ")


            except (ValueError, TypeError, IndexError):
                if statusStatus[0] != "NA":
                    indiceStatus = dados.index[dados['Código'] == produtos].tolist()
                    linhaStatus = indiceStatus[0] + 2
                    planilhaEstoque.update_cell(row=int(linhaStatus), col=14, value="NA")
                    planilhaEstoque.update_cell(row=int(linhaStatus), col=15, value="NA")

        # Segunda condição: CORREÇÃO - lógica mais simples
        if (esStatus[0] != " - " and 
            ordemdeCompraStatus[0] not in [" - ", "NA", ""] and
            statusStatus[0] != "PO Aberta"):
            indiceStatus = dados.index[dados['Código'] == produtos].tolist()
            linhaStatus = indiceStatus[0] + 2
            planilhaEstoque.update_cell(row=int(linhaStatus), col=14, value="PO Aberta")

        # Terceira condição: Se for semiacabado
        if esStatus[0] == " - " and statusStatus[0] != "NA":
            indiceStatus = dados.index[dados['Código'] == produtos].tolist()
            linhaStatus = indiceStatus[0] + 2
            planilhaEstoque.update_cell(row=int(linhaStatus), col=14, value="NA")
            planilhaEstoque.update_cell(row=int(linhaStatus), col=15, value="NA")

def gerar_pdf_tabela_multipagina(titulo="ESTOQUE", nome_arquivo="tabela_estoque.pdf", max_linhas_por_pagina=35):
    """
    Gera um PDF com tabela que quebra automaticamente em múltiplas páginas
    
    Parâmetros:
    - dados: DataFrame ou lista de listas com os dados
    - titulo: Título do documento
    - nome_arquivo: Nome do arquivo de saída
    - max_linhas_por_pagina: Número máximo de linhas por página
    """

    dadospdf = dados[["Código", "Nome", "Família", "Categoria", "Tamanho", "Localização", "Unidade", "Quantidade Atual"]]

    dadospdf["Inventário"] = ""

    dadospdf = dadospdf.rename(columns={
        'Código': 'Cod', 
        'Nome': 'Nome', 
        'Família': 'Família',
        'Categoria': 'Categoria',
        'Localização': 'Loc',
        'Tamanho': 'T',
        'Unidade': 'Unidade',
        'Quantidade Atual': 'Quan',
        'Inventário': 'I'
    })
    
    # Configurar PDF em landscape
    pdf = FPDF(orientation="portrait", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # LARGURAS PERSONALIZADAS PARA CADA COLUNA
    larguras_personalizadas = {
        'Cod': 10,      
        'Nome': 35,         
        'Família': 25, 
        'T': 10, 
        'Categoria': 30,     
        'Loc': 20,   
        'Unidade': 33,
        'Quan': 13,  
        'I': 13
    }
    
    # Ordem das colunas (a mesma do DataFrame)
    colunas_ordenadas = ['Cod', 'Nome', 'Família', 'Categoria', 'T', 'Loc', 'Unidade', 
                         'Quan', 'I']
    
    # Verificar se a soma das larguras cabe na página
    largura_pagina = 280  # A4 landscape width
    margens = 10
    largura_util = largura_pagina - (2 * margens)
    largura_total = sum(larguras_personalizadas.values())
    
    # Ajustar proporcionalmente se ultrapassar a largura útil
    if largura_total > largura_util:
        fator_ajuste = largura_util / largura_total
        for coluna in larguras_personalizadas:
            larguras_personalizadas[coluna] *= fator_ajuste
    
    # Processar dados em chunks
    total_linhas = len(dadospdf)
    num_paginas = (total_linhas + max_linhas_por_pagina - 1) // max_linhas_por_pagina
    
    for pagina in range(num_paginas):
        pdf.add_page()
        
        # Cabeçalho da página
        pdf.set_font("Arial", style="B", size=14)
        pdf.cell(0, 10, titulo, ln=True, align=Align.C)
        
        # Número da página
        pdf.set_font("Arial", style="I", size=8)
        pdf.cell(0, 5, f"Página {pagina + 1} de {num_paginas}", ln=True, align=Align.R)
        
        # Calcular range de linhas para esta página
        inicio = pagina * max_linhas_por_pagina
        fim = min((pagina + 1) * max_linhas_por_pagina, total_linhas)
        dadospdf_pagina = dadospdf.iloc[inicio:fim]
        
        # Cabeçalho da tabela
        pdf.set_font("Arial", style="B", size=10)
        pdf.ln(5)
        
        # Desenhar linha do cabeçalho com larguras personalizadas
        pdf.set_fill_color(200, 200, 200)
        for coluna in colunas_ordenadas:
            largura = larguras_personalizadas[coluna]
            pdf.cell(largura, 8, str(coluna), border=1, fill=True, align=Align.C)
        pdf.ln()
        
        # Dados da tabela com larguras personalizadas
        pdf.set_font("Arial", size=9)
        for indice, linha in dadospdf_pagina.iterrows():
            # Alternar cores para melhor legibilidade
            if indice % 2 == 0:
                pdf.set_fill_color(240, 240, 240)
            else:
                pdf.set_fill_color(255, 255, 255)
            
            for coluna in colunas_ordenadas:
                largura = larguras_personalizadas[coluna]
                valor = linha[coluna]
                
                # Truncar texto muito longo (especialmente para coluna Nome)
                texto = str(valor)
                if coluna == 'Nome' and len(texto) > 35:
                    texto = texto[:32] + "..."
                elif len(texto) > 20:
                    texto = texto[:17] + "..."
                
                
                alinhamento = Align.C
                
                pdf.cell(largura, 6, texto, border=1, fill=True, align=alinhamento)
            pdf.ln()
    
    # Salvar arquivo

    arquivo_salvo = salvar_pdf_no_drive(
                        pdf=pdf,
                        nome_arquivo="estoque_pdf",
                        pasta_id=st.secrets["id_estoque"]
                    )
    if arquivo_salvo:
        st.success("PDF salvo com sucesso no Google Drive!")
        st.write(f"**Nome:** {arquivo_salvo['name']}")
        if arquivo_salvo.get('webViewLink'):
            st.write(f"**Link:** {arquivo_salvo['webViewLink']}")

    return None

status()

with st.sidebar:
    selected = option_menu(
        menu_title="Menu",
        options=["Painel de Controle", "Controle de Inventário", "Apontamento", "Cadastro de Produtos", "Ordem de Compra", "Edição de Informações"],
        icons = ["bar-chart", "box-seam", "bi-arrow-up-right-square" ,"plus-circle", "bag-plus", "pencil-square"],
        menu_icon="cast",
        default_index=0,
    )

    if st.button("Gerar pdf do Estoque"):
        gerar_pdf_tabela_multipagina()

    # if st.button("Mandar email", width="stretch"):
    #     mandar_email()

if selected == "Painel de Controle":
    # CSS personalizado que se adapta ao tema
    st.markdown("""
        <style>
        /* Estilos que se adaptam ao tema */
        .company-header {
            font-size: 3.5rem;
            font-weight: 700;
            text-align: center;
            margin-bottom: 0.5rem;
            padding: 2rem;
            background-image: url('Estamparia2.png');
            background-size: cover;
            background-position: center;
            border-radius: 10px;
            color: white;
            text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.7);
        }
        .company-tagline {
            font-size: 1.5rem;
            text-align: center;
            margin-bottom: 2rem;
        }
        .section-header {
            font-size: 2rem;
            border-bottom: 2px solid;
            padding-bottom: 0.5rem;
            margin: 2rem 0 1.5rem 0;
        }
        .service-card {
            border-radius: 10px;
            padding: 1.5rem;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            height: 100%;
            transition: transform 0.3s;
        }
        .service-card:hover {
            transform: translateY(-5px);
        }
        .contact-info {
            padding: 2rem;
            border-radius: 10px;
            margin-top: 2rem;
        }
        .footer {
            text-align: center;
            padding: 1.5rem;
            margin-top: 3rem;
            font-size: 0.9rem;
        }
        
        /* Ajustes específicos para tema claro */
        [data-theme="light"] {
            --text-color: #31333F;
            --bg-color: #f8f9fa;
            --card-bg: white;
            --primary-color: #1a73e8;
            --secondary-color: #5f6368;
            --gradient: linear-gradient(to right, #ffffff, #f1f3f5);
        }
        
        /* Ajustes específicos para tema escuro */
        [data-theme="dark"] {
            --text-color: #FFFFFF;
            --bg-color: #0E1117;
            --card-bg: #262730;
            --primary-color: #3eb0f7;
            --secondary-color: #AAAAAA;
            --gradient: linear-gradient(to right, #0E1117, #1a1a2e);
        }
        
        /* Aplicação das variáveis CSS */
        .main {
            background-color: var(--bg-color);
        }
        .stApp {
            background: var(--gradient);
        }
        .company-tagline {
            color: var(--secondary-color);
        }
        .section-header {
            color: var(--primary-color);
            border-bottom-color: var(--primary-color);
        }
        .service-card {
            background-color: var(--card-bg);
            color: var(--text-color);
        }
        .contact-info {
            background-color: var(--primary-color);
            color: white;
        }
        .footer {
            color: var(--secondary-color);
        }
        </style>
    """, unsafe_allow_html=True)

    # Header da empresa com imagem de fundo
    #st.markdown('<h1 class="company-header">ATLETA PRO</h1>', unsafe_allow_html=True)

    # Imagem de destaque
    co1, co2, co3 = st.columns((1, 1, 1))

    with co1:
        st.image("Imagem1.png", use_container_width=True)

    with co2:
        st.markdown(
            """
            <div style="text-align: center">
                <h3>ATLETA PRO</h3>
            </div>
            """, 
            unsafe_allow_html=True
        )
        st.image("Imagem3.png", use_container_width=True)

    with co3:
        st.image("Imagem2.png", use_container_width=True)

    # Seção de serviços
    st.markdown('<h2 class="section-header"></h2>', unsafe_allow_html=True)


    col1, col2 = st.columns(2)

    with col1:
        atencao = dados[dados["Status"] == "Atenção"]

        status()

        if atencao.empty: 
            st.subheader(f"Nenhum produto com estoque abaixo do Estoque de Segurança")
        else:
            patencao = atencao['Nome'].tolist()
            codigoatencao = atencao['Código'].tolist()
            fornecedoratencao = atencao['Fornecedor Principal'].tolist()
            contatoatencao = atencao['Contato Fornecedor'].tolist()
            atualatencao = atencao['Quantidade Atual'].tolist()
            segurancaatencao = atencao['Estoque de Segurança'].tolist()
            minimoatencao = atencao['Lote mínimo'].tolist()
            st.subheader("Produto(s) com estoque crítico:")
            i = 0

            with st.expander("Mostrar Produtos com Estoque Crítico"):
                for p in patencao:
                    st.divider()    
                    st.subheader(f"\n- Produto: {p}.\n")
                    st.markdown(f"Código: {codigoatencao[i]}.\n")
                    st.markdown(f"Fornecedor: {fornecedoratencao[i]}.\n")
                    st.markdown(f"Contato: {contatoatencao[i]}.\n")
                    st.markdown(f"Quantidade atual: {atualatencao[i]}.\n")
                    st.markdown(f"Estoque de Segurança: {segurancaatencao[i]}.\n")
                    st.markdown(f"Lote mínimo: {minimoatencao[i]}.\n")
                    st.subheader(f"ABRA ORDEM DE COMPRA!\n")
                    i += 1
                         


    with col2:
        aberto = dados[dados["Status"] == "PO Aberta"]

        if aberto.empty: 
            st.subheader(f"Nenhum produto com Ordem de Compra Aberta")
        else:
            paberto = aberto['Nome'].tolist()
            codigoaberto = aberto['Código'].tolist()
            fornecedoraberto = aberto['Fornecedor Principal'].tolist()
            contatoaberto = aberto['Contato Fornecedor'].tolist()
            quantidadeaberto = aberto['Quantidade Atual'].tolist()
            segurancaaberto = aberto['Estoque de Segurança'].tolist()
            comproaberto= aberto['PO'].tolist()
            previsaoPO= aberto['Previsão PO'].tolist()
            st.subheader("Produto(s) com Ordem de Compra Aberta:")
            i = 0

            with st.expander("Mostrar Produtos com Ordem de Compra Aberta"):
                for p in paberto:
                    st.divider()
                    st.subheader(f"\nProduto: {p}.\n")
                    st.markdown(f"Código: {codigoaberto[i]}.\n")
                    st.markdown(f"Fornecedor: {fornecedoraberto[i]}.\n")
                    st.markdown(f"Contato: {contatoaberto[i]}.\n")
                    st.markdown(f"Quantidade Atual: {quantidadeaberto[i]}.\n")
                    st.markdown(f"Estoque de Segurança: {segurancaaberto[i]}.\n")
                    st.markdown(f"Comprado: {comproaberto[i]}.\n")
                    st.markdown(f"Previsão de chegada: {previsaoPO[i]}.\n")
                    if st.button(f"Fechar Ordem de Compra", key=f"btn_{codigoaberto[i]}"):
                        indicePO = dados.index[dados['Nome'] == p].tolist()
                        linhaPO = indicePO[0] + 2
                        planilhaEstoque.batch_update([{
                            'range': f'K{linhaPO}:L{linhaPO}',  # Colunas A até J (1 a 10)
                            'values': [[" - ", " - "]]
                        }])  
                        st.toast("Lembre-se de fazer o Apontamento do produto pedido nessa Ordem de Compra")
                        st.rerun()
                    i += 1

    ##Plotando dados falsos ##
    st.markdown('<h2 class="section-header"></h2>', unsafe_allow_html=True)

    # Gerar os dados
    #df_producao = gerar_dados_producao() 
    
    
    familia = dados["Família"].unique().tolist()   
    unique_families = dados["Família"].unique()
    color_palette = px.colors.qualitative.Dark2
    family_colors = {family: color_palette[i % len(color_palette)] 
                    for i, family in enumerate(unique_families)}

    # Primeiro, vamos agrupar por Família e Categoria
    grupos = dados.groupby(['Família', 'Categoria'])

    for (f, categoria), db_grupo in grupos:
        st.subheader(f"Gráfico da família {f} - Categoria {categoria}:")
        
        # Ordenar por Tamanho se existir a coluna
        if 'Tamanho' in db_grupo.columns:
            db_grupo = db_grupo.sort_values('Tamanho')
        
        # Criar um label para o eixo Y que inclui Tamanho se existir
        if 'Tamanho' in db_grupo.columns:
            db_grupo['Label_Y'] = db_grupo['Nome'] + ' (' + db_grupo['Tamanho'].astype(str) + ')'
        else:
            db_grupo['Label_Y'] = db_grupo['Nome']
        
        # Converter estoque de segurança para numérico onde possível
        db_grupo['Estoque_Seguranca_Num'] = pd.to_numeric(db_grupo['Estoque de Segurança'], errors='coerce')
        
        db_grupo['StatusE'] = db_grupo.apply(lambda row: 
            'Abaixo do Estoque' if (pd.notna(row['Estoque_Seguranca_Num']) and 
                                row['Quantidade Atual'] < row['Estoque_Seguranca_Num']) 
            else 'Acima do Estoque' if pd.notna(row['Estoque_Seguranca_Num'])
            else 'Outro', axis=1)

        # Definir as cores para cada status
        status_colors = {
            'Abaixo do Estoque': '#FF6B6B',  # Vermelho suave
            'Acima do Estoque': '#51CF66',   # Verde esmeralda
            'Outro': '#868E96'               # Cinza azulado
        }

        st.write(db_grupo['StatusE'].iloc[0])
        if db_grupo['StatusE'].iloc[0] == "Outro":
            fig = px.bar(db_grupo, 
                                y='Label_Y', 
                                x="Quantidade Atual", 
                                color="Tamanho",
                                color_discrete_sequence=px.colors.qualitative.Prism,
                                title=f'Família {f} - Categoria {categoria}', 
                                orientation='h',
                                text='Quantidade Atual',
                                hover_data={
                                    'Nome': True,
                                    'Tamanho': True if 'Tamanho' in db_grupo.columns else False,
                                    'Quantidade Atual': ':.0f',
                                    'Estoque de Segurança': True,
                                    'StatusE': False,
                                    'Família': False,
                                    'Categoria': False,
                                    'Label_Y': False
                                })
                
            # Personalização do layout
            fig.update_layout(
                title={
                    'text': f"Família {f} - Categoria {categoria}",
                    'y': 0.95,
                    'x': 0.5,
                    'xanchor': 'center',
                    'yanchor': 'top',
                    'font': {'size': 20}
                },
                xaxis_title='Quantidade',
                yaxis_title='Produto (Tamanho)' if 'Tamanho' in db_grupo.columns else 'Produto',
                height=max(400, len(db_grupo) * 30),  # Altura dinâmica baseada no número de produtos
                hovermode='y unified',
                showlegend=True,
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.02,
                    xanchor="right",
                    x=1,
                    title_text='Status do Estoque:'
                ),
                margin=dict(l=200, r=50, t=100, b=50),  # Aumentei margem esquerda para labels maiores
                uniformtext_minsize=10,
                uniformtext_mode='hide',
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)'
            )

        else:
            fig = px.bar(db_grupo, 
                                y='Label_Y', 
                                x="Quantidade Atual", 
                                color="StatusE",
                                color_discrete_map=status_colors,
                                title=f'Família {f}', 
                                orientation='h',
                                text='Quantidade Atual',
                                hover_data={
                                    'Nome': True,
                                    'Quantidade Atual': ':.0f',
                                    'Estoque de Segurança': True,
                                    'StatusE': False,
                                    'Categoria': False,
                                    'Label_Y': False
                                })
                
        # Personalização do layout
        fig.update_layout(
            title={
                'text': f"Família {f}",
                'y': 0.95,
                'x': 0.5,
                'xanchor': 'center',
                'yanchor': 'top',
                'font': {'size': 20}
            },
            xaxis_title='Quantidade',
            yaxis_title='Produto (Tamanho)' if 'Tamanho' in db_grupo.columns else 'Produto',
            height=max(400, len(db_grupo) * 30),  # Altura dinâmica baseada no número de produtos
            hovermode='y unified',
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1,
                title_text='Status do Estoque:'
            ),
            margin=dict(l=200, r=50, t=100, b=50),  # Aumentei margem esquerda para labels maiores
            uniformtext_minsize=10,
            uniformtext_mode='hide',
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)'
        )

        # Personalização do hover
        hovertemplate = '<b>%{customdata[0]}</b><br>'
        if 'Tamanho' in db_grupo.columns:
            hovertemplate += 'Tamanho: %{customdata[1]}<br>'
        hovertemplate += 'Quantidade Atual: %{x:.0f}<br>' + \
                        'Estoque Segurança: %{customdata[2 if "Tamanho" in db_grupo.columns else 1]}<br>' + \
                        '<extra></extra>'

        fig.update_traces(
            texttemplate='%{text:.0f}',
            textposition='outside',
            hovertemplate=hovertemplate
        )

        # Adicionar linhas verticais APENAS para produtos com estoque numérico
        db_com_estoque = db_grupo[pd.notna(db_grupo['Estoque_Seguranca_Num'])]
        if not db_com_estoque.empty:
            fig.add_trace(go.Scatter(
                y=db_com_estoque['Label_Y'],
                x=db_com_estoque['Estoque_Seguranca_Num'],
                name='Estoque de Segurança',
                mode='markers',
                marker=dict(
                    color='lightgray',
                    size=10,
                    symbol='line-ns-open',
                    line=dict(width=2, color='red')
                ),
                hoverinfo='y+x',
                hovertemplate='<b>%{y}</b><br>Estoque Segurança: %{x:.0f}<extra></extra>'
            ))

        st.plotly_chart(fig, use_container_width=True)
        st.divider()
    
elif selected == "Controle de Inventário":
    st.title("Controle de Inventário")

    st.write('- - -')

    opcoes = dados["Família"].unique().tolist()
    escolha = st.selectbox("Escolha uma família de produtos", opcoes)
    
    familia = dados[dados["Família"] == escolha]

    categorias = familia['Categoria'].unique().tolist()
    
    categoria = st.selectbox("Escolha uma categoria de produto:", categorias)
    
    # Encontrar a linha correspondente ao produto selecionado
    categoria = familia[familia['Categoria'] == categoria]

    tamanhos = categoria['Tamanho'].unique().tolist()
    
    tamanho = st.selectbox("Escolha o tamanho:", tamanhos)

    tamanho = categoria[categoria['Tamanho'] == tamanho]

    nomes = tamanho['Nome'].tolist()
    
    produto_selecionado = st.radio("ESCOLHA UM PRODUTO:", nomes)
    
    # Encontrar a linha correspondente ao produto selecionado
    linha_produto = tamanho[tamanho['Nome'] == produto_selecionado]

    st.write(linha_produto)

    if not linha_produto.empty:
        quantidadeRegistrada = linha_produto['Quantidade Atual'].values[0] 
        unidadeRegistrada = linha_produto['Unidade'].values[0] 
        indice = dados.index[dados['Código'] == linha_produto['Código'].iloc[0]].tolist()
        linha = indice[0] + 2 if indice else st.write("Problema")

        st.write(f"Última quantidade registrada: {quantidadeRegistrada}")
        st.write(f"Unidade registrada: {unidadeRegistrada}")

        with st.form("Ajuste_Estoque"):
            quantidade = st.text_input(f"Quantidade de {produto_selecionado} atual:")
            # Botão de submit dentro do form
            submitted = st.form_submit_button("Registrar")
            
            # Só executa quando o botão for clicado
            if submitted:
                planilhaEstoque.update_cell(row=int(linha),col=10,value=str(quantidade))    
                status()
                st.rerun()
                st.toast("Quantidade atualizada!")

    else:
        st.error("Produto não encontrado!")

elif selected == "Apontamento":
    st.title("Apontamento")

    InputOutput = st.pills("Deseja registrar Entrada ou Saída de um estoque", ["Entrada", "Saída"])

    st.write('- - -')

    opcoes = dados["Categoria"].unique().tolist()
    escolha = st.selectbox("Escolha uma família de produtos", opcoes)
    
    categoria = dados[dados["Categoria"] == escolha]

    # Obter lista de nomes
    nomes = categoria['Nome'].tolist()
    
    produto_selecionado = st.radio("ESCOLHA UM PRODUTO:", nomes)
    
    # Encontrar a linha correspondente ao produto selecionado
    linha_produto = categoria[categoria['Nome'] == produto_selecionado]

    if not linha_produto.empty:
        quantidadeRegistrada = linha_produto['Quantidade Atual'].values[0]
        localRegistrado = linha_produto['Localização'].values[0] 
        unidadeRegistrada = linha_produto['Unidade'].values[0] 
        indice = dados.index[dados['Nome'] == produto_selecionado].tolist()
        linha = indice[0] + 2

        st.write(f"Localização: {localRegistrado}")
        st.write(f"Última quantidade registrada: {quantidadeRegistrada}")
        st.write(f"Unidade registrada: {unidadeRegistrada}")
        
        quantidade = st.number_input(f"Quantidade de {produto_selecionado} em movimento ({InputOutput}):", step=1)

        if InputOutput == "Entrada":
            quantidade = (quantidade) + (quantidadeRegistrada)
            # Botão de submit dentro do form
            submitted = st.button("Submeter")
            
            # Só executa quando o botão for clicado
            if submitted:
                planilhaEstoque.update_cell(row=int(linha),col=10,value=int(quantidade))
                st.toast(f"Entrada de {produto_selecionado} registrada")
                status()
                st.rerun()

        elif InputOutput == "Saída":
            quantidade = (quantidadeRegistrada) - (quantidade)

            if quantidade < 0:
                st.error("Estoque incompatível!")

            else:
                # Botão de submit dentro do form
                submitted = st.button("Submeter")
                
                # Só executa quando o botão for clicado
                if submitted:
                    planilhaEstoque.update_cell(row=int(linha),col=10,value=int(quantidade)) 
                    status()
                    st.toast(f"Saída de {produto_selecionado} registrada")
                    st.rerun()

        else:
            st.error("Selecione tipo de movimentação!")
        
    
    else:
        st.error("Produto não encontrado!")

elif selected == "Cadastro de Produtos":
    st.title("Cadastro de Produtos")

    st.write("- - -")

    insumoOrSemiacabado = st.pills("Deseja registrar Insumo ou Semiacabado", ["Insumo", "Semiacabado"])

    # Verifica se o DataFrame existe E tem dados
    if dados is not None and len(dados) > 0:
        ultima_linha = dados.index[-1] + 3
    else:
        ultima_linha = 2  # primeira linha de dados

    if insumoOrSemiacabado ==  "Insumo":
        with st.form("meu_formulario"):
            codigoProduto = st.text_input("Informe o código do insumo", max_chars= 50)

            nomeProduto = st.text_input("Informe o nome do insumo", max_chars= 50)

            familiaProduto = st.selectbox("Selecione a família do insumo dos tipos já registrados ou ESCREVA UMA NOVA", dados["Família"].unique(), accept_new_options=True)

            fornecedorProduto = st.text_input("Informe o fornecedor do insumo", max_chars= 50)

            contatoFornecedor = st.text_input("Informe o contato do fornecedor do insumo", max_chars= 50)

            localizacaoEstoque = st.text_input("Informe a localização do insumo no estoque", max_chars= 50)

            unidadeProduto = st.text_input("Informe o tipo de unidade do insumo", max_chars= 50)

            quantidadeProduto = st.number_input("Informe a quantidade do insumo de acordo com a unidade registrada", step=1)

            esProduto = st.number_input("Informe a quantidade de estoque de segurança para o insumo na unidade registrada", step=1)

            loteminimoProduto = st.number_input("Informe o lote mínimo em unidades registradas", step=1)

            observacaoProduto = st.text_input("Observações:", max_chars= 70)
            
            # Botão de submit dentro do form
            submitted = st.form_submit_button("Submeter")
            
            # Só executa quando o botão for clicado
            if submitted:
                # Criar a lista de valores para a linha inteira
                valores_linha = [
                    codigoProduto,
                    nomeProduto,
                    familiaProduto,
                    "-",
                    "-",
                    fornecedorProduto,
                    contatoFornecedor,
                    localizacaoEstoque,
                    unidadeProduto,
                    quantidadeProduto,  # Convertendo para int como no exemplo original
                    esProduto,
                    loteminimoProduto,
                    observacaoProduto
                ]

                # Fazer o batch update
                planilhaEstoque.batch_update([{
                    'range': f'A{ultima_linha}:M{ultima_linha}',  # Colunas A até J (1 a 10)
                    'values': [valores_linha]
                }])
                

                status()

                st.toast("Insumo adicionado com sucesso!", icon="🎉")


    elif insumoOrSemiacabado ==  "Semiacabado":
        with st.form("meu_formulario"):
            codigoProduto = st.text_input("Informe o código do semiacabado", max_chars= 50)

            nomeProduto = st.text_input("Informe o nome do semiacabado", max_chars= 50)

            familiaProduto = st.selectbox("Selecione a família do semiacabado dos tipos já registrados ou ESCREVA UM NOVO TIPO", dados["Família"].unique(), accept_new_options=True)

            categoriaProduto = st.selectbox("Selecione a categoria do semiacabado dos tipos já registrados ou ESCREVA UM NOVO TIPO", dados["Categoria"].unique(), accept_new_options=True)

            tamanhoProduto = st.selectbox("Selecione o tamanho do semiacabado dos tipos já registrados ou ESCREVA UM NOVO TIPO", dados["Tamanho"].unique(), accept_new_options=True)

            fornecedorProduto = " - "

            contatoFornecedor = " - "

            localizacaoEstoque = st.text_input("Informe a localização do semiacabado no estoque", max_chars= 50)

            unidadeProduto = st.text_input("Informe o tipo de unidade do semiacabado", max_chars= 50)

            quantidadeProduto = st.number_input("Informe a quantidade do semiacabado de acordo com a unidade registrada", step=1)

            esProduto = " - "

            loteminimoProduto = " - "

            observacaoProduto = st.text_input("Observações:", max_chars= 70)

            statusProduto = "NA"

            po = "NA"
            
            # Botão de submit dentro do form
            submitted = st.form_submit_button("Submeter")
            
            # Só executa quando o botão for clicado
            if submitted:
                # Criar a lista de valores para a linha inteira
                valores_linha = [
                    codigoProduto,
                    nomeProduto,
                    familiaProduto,
                    categoriaProduto,
                    tamanhoProduto,
                    fornecedorProduto,
                    contatoFornecedor,
                    localizacaoEstoque,
                    unidadeProduto,
                    quantidadeProduto,  # Convertendo para int como no exemplo original
                    esProduto,
                    loteminimoProduto,
                    observacaoProduto,
                    statusProduto,
                    po
                ]

                # Fazer o batch update
                planilhaEstoque.batch_update([{
                    'range': f'A{ultima_linha}:O{ultima_linha}',  # Colunas A até J (1 a 10)
                    'values': [valores_linha]
                }])

                status()

                st.toast("Semiacabado adicionado com sucesso!", icon="🎉")

elif selected == "Ordem de Compra":
    st.title("Abrir Ordem de Compra")

    st.write("- - -")

    insumoPO = dados[dados["Estoque de Segurança"] != " - "]

    opcoesPO = insumoPO["Categoria"].unique().tolist()
    escolhaPO = st.selectbox("Escolha uma família de produtos", opcoesPO)
    
    categoriaPO = insumoPO[insumoPO["Categoria"] == escolhaPO]

    # Obter lista de nomes
    nomesPO = categoriaPO['Nome'].tolist()
    
    produto_selecionadoPO = st.selectbox("ESCOLHA UM PRODUTO:", nomesPO)

    st.write("- - -")

    unidadePO = insumoPO["Unidade"].unique().tolist()

    linha_produtoPO = categoriaPO[categoriaPO['Nome'] == produto_selecionadoPO]

    ordemPO = linha_produtoPO["PO"].tolist()

    with st.form("abrirPO"):

        st.write(f"Quantidade atual: {linha_produtoPO['Quantidade Atual'].iloc[0]}")

        if ordemPO[0] != " - ":
            st.error("Produto já possui ordem de compra em aberto!")
            st.warning("Em caso de nova PO, informe na quantidade a soma de todos os pedidos em aberto!")

        quantidadeProdutoPO = st.number_input(f"Informe a quantidade de {produto_selecionadoPO} em {linha_produtoPO['Unidade'].iloc[0]}", step=1,value=linha_produtoPO["Lote mínimo"].iloc[0])

        previsaoPO = st.date_input(f"Informe a previsão de chegada da ordem de serviço de {produto_selecionadoPO}.", value='today', format="DD/MM/YYYY")
        previsaoPO = previsaoPO.strftime("%d/%m/%Y")

        submitted = st.form_submit_button("Abrir PO")

        indicePO = dados.index[dados['Nome'] == produto_selecionadoPO].tolist()
        linhaPO = indicePO[0] + 2

        if submitted:

            planilhaEstoque.update_cell(row=int(linhaPO), col=14, value="PO Aberta")
            planilhaEstoque.update_cell(row=int(linhaPO), col=15, value=quantidadeProdutoPO)
            planilhaEstoque.update_cell(row=int(linhaPO), col=16, value=previsaoPO)

            st.toast("Ordem de Compra registrada!", icon="🎉")

            st.rerun()          
            
elif selected == "Edição de Informações":
    st.title("Edição de Informações")

    st.write("- - -")

    opcoesEdicao = dados["Categoria"].unique().tolist()
    escolhaEdicao = st.selectbox("Escolha uma família de produtos", opcoesEdicao)
    
    categoriaEdicao = dados[dados["Categoria"] == escolhaEdicao]

    # Obter lista de nomes
    nomesEdicao = categoriaEdicao['Nome'].tolist()
    
    produto_selecionadoEdicao = st.selectbox("ESCOLHA UM PRODUTO:", nomesEdicao)

    st.write("- - -")

    st.header("Edite as informações:")

    with st.form("editeInfo"):

        linha_produtoEdicao = categoriaEdicao[categoriaEdicao['Nome'] == produto_selecionadoEdicao]

        codigoProdutoEdicao = st.text_input("Informe o novo código do produto", max_chars= 50, value=linha_produtoEdicao["Código"].iloc[0])

        nomeProdutoEdicao = st.text_input("Informe o nome do produto", max_chars= 50, value=linha_produtoEdicao["Nome"].iloc[0])

        categoriaProdutoEdicao = st.text_input("Edite a categoria do produto", max_chars= 50, value=linha_produtoEdicao["Categoria"].iloc[0])

        fornecedorProdutoEdicao = st.text_input("Informe o fornecedor do produto", max_chars= 50, value=linha_produtoEdicao["Fornecedor Principal"].iloc[0])

        contatoFornecedorEdicao = st.text_input("Informe o contato do fornecedor do produto", max_chars= 50, value=linha_produtoEdicao["Contato Fornecedor"].iloc[0])

        localizacaoEstoqueEdicao = st.text_input("Informe a localização do produto no estoque", max_chars= 50, value=linha_produtoEdicao["Localização"].iloc[0])

        unidadeProdutoEdicao = st.text_input("Informe o tipo de unidade do produto", max_chars= 50, value=linha_produtoEdicao["Unidade"].iloc[0])

        quantidadeProdutoEdicao = st.number_input("Informe a quantidade do produto de acordo com a unidade registrada", step=1, value=linha_produtoEdicao["Quantidade Atual"].iloc[0])

        loteMinProdutoEdicao = st.number_input("Informe o lote mínimo em unidades registradas", step=1, value=linha_produtoEdicao["Lote mínimo"].iloc[0])

        esProdutoEdicao = st.number_input("Informe a quantidade de estoque de segurança para o produto na unidade registrada", step=1, value=linha_produtoEdicao["Estoque de Segurança"].iloc[0])

        submitted = st.form_submit_button("Editar")

        if submitted:
            indiceEdicao = dados.index[dados['Nome'] == produto_selecionadoEdicao].tolist()
            linhaEdicao = indiceEdicao[0] + 2


            # Criar a lista de valores para a linha inteira
            valores_linha = [
                codigoProdutoEdicao,
                nomeProdutoEdicao,
                categoriaProdutoEdicao,
                fornecedorProdutoEdicao,
                contatoFornecedorEdicao,
                localizacaoEstoqueEdicao,
                unidadeProdutoEdicao,
                int(quantidadeProdutoEdicao),  # Convertendo para int como no exemplo original
                esProdutoEdicao,
                loteMinProdutoEdicao
            ]

            # Fazer o batch update
            planilhaEstoque.batch_update([{
                'range': f'A{linhaEdicao}:J{linhaEdicao}',  # Colunas A até J (1 a 10)
                'values': [valores_linha]
            }])

            status()

            st.toast("Produto adicionado com sucesso!", icon="🎉")

            st.rerun()