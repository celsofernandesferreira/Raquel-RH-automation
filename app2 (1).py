import streamlit as st
import google.generativeai as genai
import json
import re
import io
import docx
import logging
from datetime import datetime
import openpyxl
import pandas as pd

# ============================================================
# 1. CONFIGURAÇÃO DE LOGS E PÁGINA
# ============================================================
logging.basicConfig(
    filename="automacao_rh_raquel.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)

st.set_page_config(page_title="Automação de Relatórios RH", page_icon="📊", layout="wide")
st.title("📊 Sistema de Automação de Relatórios — RH")
st.subheader("Extração Inteligente de Notas para Excel")

# ============================================================
# 2. INICIALIZAÇÃO DA API DO GEMINI
# ============================================================
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
except Exception:
    st.error("Erro: Chave API em falta nos Secrets do Streamlit.")
    st.stop()

MODELO_GEMINI = "gemini-2.5-flash"

# ============================================================
# 3. ESTADO DA SESSÃO
# ============================================================
if "dados_extraidos" not in st.session_state:
    st.session_state.dados_extraidos = None
if "excel_pronto" not in st.session_state:
    st.session_state.excel_pronto = None
if "nome_ficheiro_saida" not in st.session_state:
    st.session_state.nome_ficheiro_saida = None

# ============================================================
# 4. FUNÇÕES DE LEITURA
# ============================================================
def ler_ficheiro_txt(uploaded_file):
    return uploaded_file.read().decode("utf-8", errors="ignore")

def ler_ficheiro_docx(uploaded_file):
    doc = docx.Document(io.BytesIO(uploaded_file.read()))
    return "\n".join([para.text for para in doc.paragraphs])

# ============================================================
# 5. DESCOBERTA AUTOMÁTICA DO CABEÇALHO NO EXCEL
# ============================================================
def extrair_cabecalho_automatico(excel_bytes):
    """Lê o Excel e tenta adivinhar qual é a linha de cabeçalho (a que tem mais texto)."""
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    ws = wb.active
    
    melhor_linha = 1
    max_colunas_preenchidas = 0
    mapa_colunas = {} # Nome exato da coluna no Excel -> Número da coluna

    # Procura nas primeiras 30 linhas
    for r in range(1, min(30, ws.max_row) + 1):
        mapa_temp = {}
        for c in range(1, ws.max_column + 1):
            val = ws.cell(row=r, column=c).value
            if val and isinstance(val, str) and val.strip():
                mapa_temp[val.strip()] = c
                
        # Assumimos que a linha com mais células de texto é o cabeçalho
        if len(mapa_temp) > max_colunas_preenchidas:
            max_colunas_preenchidas = len(mapa_temp)
            melhor_linha = r
            mapa_colunas = mapa_temp

    return melhor_linha, mapa_colunas

# ============================================================
# 6. AGENTE DE INTELIGÊNCIA ARTIFICIAL (AGORA INTELIGENTE)
# ============================================================
def extrair_dados_com_gemini(texto_notas, nomes_colunas_excel):
    model = genai.GenerativeModel(model_name=MODELO_GEMINI)

    prompt_sistema = f"""
    Tu és um Assistente de RH Avançado especialista em extração de dados.
    O teu objetivo é ler as notas em texto livre e extrair a informação, compreendendo o seu contexto.

    O ficheiro Excel de destino tem EXATAMENTE estas colunas:
    {nomes_colunas_excel}

    Regras:
    1. Devolve APENAS um array JSON com objetos. Cada objeto é um registo/pessoa diferente.
    2. Usa EXATAMENTE os nomes das colunas fornecidas acima como as chaves do teu JSON.
    3. Usa a tua inteligência semântica para decidir qual informação do texto pertence a qual coluna.
    4. Se a informação para uma determinada coluna não constar nas notas, coloca o valor como "".
    5. Não adiciones texto antes ou depois do JSON.
    """

    try:
        response = model.generate_content(
            [prompt_sistema, f"Notas para analisar:\n\n{texto_notas}"],
            generation_config={"temperature": 0.1},
        )

        texto_resposta = (response.text or "").strip()
        if "```" in texto_resposta:
            texto_resposta = re.sub(r"```json|```", "", texto_resposta).strip()

        match = re.search(r"\[.*\]", texto_resposta, re.DOTALL)
        if match:
            texto_resposta = match.group(0)

        return json.loads(texto_resposta)
    except Exception as e:
        logging.error(f"Erro IA: {e}")
        st.error(f"Erro na extração: {e}")
        return None

# ============================================================
# 7. ESCRITA NO EXCEL
# ============================================================
def preencher_excel(excel_bytes, lista_dados, linha_cabecalho, mapa_colunas):
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    ws = wb.active

    # Descobre a próxima linha livre abaixo do cabeçalho
    proxima_linha = linha_cabecalho + 1
    while any(ws.cell(row=proxima_linha, column=c).value is not None for c in mapa_colunas.values()):
        proxima_linha += 1

    linhas_inseridas = 0
    for registo in lista_dados:
        for nome_coluna, num_coluna in mapa_colunas.items():
            # A IA envia os dados já mapeados para o nome correto
            valor = registo.get(nome_coluna, "")
            cell = ws.cell(row=proxima_linha, column=num_coluna)
            if type(cell).__name__ != "MergedCell":
                cell.value = valor
        
        proxima_linha += 1
        linhas_inseridas += 1

    output_final = io.BytesIO()
    wb.save(output_final)
    output_final.seek(0)
    return output_final, linhas_inseridas

# ============================================================
# 8. INTERFACE DO UTILIZADOR
# ============================================================
col1, col2 = st.columns([1, 1])

with col1:
    st.markdown("### 📄 1. Documentos Base (Notas)")
    ficheiros_carregados = st.file_uploader("Carrega as notas", type=["txt", "docx"], accept_multiple_files=True)

    texto_acumulado = ""
    if ficheiros_carregados:
        for f in ficheiros_carregados:
            if f.name.endswith(".txt"):
                texto_acumulado += ler_ficheiro_txt(f) + "\n"
            elif f.name.endswith(".docx"):
                texto_acumulado += ler_ficheiro_docx(f) + "\n"

    texto_colado = st.text_area("Ou cola texto aqui:", height=150)
    texto_final = (texto_acumulado + "\n" + texto_colado).strip()

with col2:
    st.markdown("### 📂 2. Excel Modelo da Empresa")
    excel_modelo = st.file_uploader("Carrega o Excel (.xlsx)", type=["xlsx"])
    
    linha_cabecalho, mapa_colunas = None, {}
    
    # Ao carregar o Excel, lê e descobre logo as colunas automaticamente!
    if excel_modelo:
        excel_bytes_lido = excel_modelo.read()
        linha_cabecalho, mapa_colunas = extrair_cabecalho_automatico(excel_bytes_lido)
        
        if mapa_colunas:
            st.success(f"✅ Cabeçalho detetado na linha {linha_cabecalho}.")
            with st.expander("Colunas detetadas no ficheiro"):
                for col_name in mapa_colunas.keys():
                    st.write(f"- {col_name}")
        else:
            st.error("❌ Não foi possível detetar colunas neste ficheiro.")

st.divider()

# ============================================================
# 9. PROCESSAMENTO
# ============================================================
if st.button("🤖 Analisar Notas de Forma Inteligente", use_container_width=True, type="primary"):
    if not texto_final:
        st.warning("Introduz notas primeiro.")
    elif not excel_modelo or not mapa_colunas:
        st.error("Carrega o Excel Modelo válido primeiro.")
    else:
        with st.spinner("A analisar o contexto das notas e a mapear para as colunas do Excel..."):
            nomes_colunas = list(mapa_colunas.keys())
            lista_dados = extrair_dados_com_gemini(texto_final, nomes_colunas)

        if lista_dados:
            st.session_state.dados_extraidos = lista_dados
            st.session_state.excel_pronto = None 
            st.success("✅ Dados extraídos com sucesso!")

# ============================================================
# 10. REVISÃO E EXPORTAÇÃO
# ============================================================
if st.session_state.dados_extraidos:
    st.markdown("### 👀 3. Confirma os dados")
    df_preview = pd.DataFrame(st.session_state.dados_extraidos)
    df_editado = st.data_editor(df_preview, use_container_width=True, num_rows="dynamic")

    if st.button("📥 Preencher e Gerar Excel", use_container_width=True, type="primary"):
        lista_dados_final = df_editado.to_dict(orient="records")
        try:
            excel_modelo.seek(0)
            output_final, linhas_inseridas = preencher_excel(
                excel_modelo.read(), lista_dados_final, linha_cabecalho, mapa_colunas
            )

            st.session_state.excel_pronto = output_final.getvalue()
            st.session_state.nome_ficheiro_saida = f"Relatorio_Final_{datetime.now().strftime('%d-%m-%Y_%H%M')}.xlsx"
            st.success(f"✨ Sucesso! {linhas_inseridas} linhas prontas.")
        except Exception as ex:
            st.error(f"Erro a gerar ficheiro: {ex}")

if st.session_state.excel_pronto:
    st.download_button(
        label=f"📥 Descarregar '{st.session_state.nome_ficheiro_saida}'",
        data=st.session_state.excel_pronto,
        file_name=st.session_state.nome_ficheiro_saida,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
