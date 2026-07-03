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
# 1. CONFIGURAÇÃO
# ============================================================
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

st.set_page_config(page_title="Automação de Formulários RH", page_icon="📝", layout="wide")
st.title("📝 Sistema de Automação — Preenchimento de Formulários de Entrevista")

try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
except Exception:
    st.error("Erro: Chave API em falta nos Secrets do Streamlit.")
    st.stop()

MODELO_GEMINI = "gemini-2.5-flash"

if "dados_extraidos" not in st.session_state:
    st.session_state.dados_extraidos = None
if "excel_pronto" not in st.session_state:
    st.session_state.excel_pronto = None

# ============================================================
# 2. FUNÇÕES DE LEITURA
# ============================================================
def ler_ficheiro_txt(uploaded_file):
    return uploaded_file.read().decode("utf-8", errors="ignore")

def ler_ficheiro_docx(uploaded_file):
    doc = docx.Document(io.BytesIO(uploaded_file.read()))
    return "\n".join([para.text for para in doc.paragraphs])

# ============================================================
# 3. MAPEAMENTO DO FORMULÁRIO (NOVA LÓGICA VERTICAL)
# ============================================================
def ler_estrutura_formulario(excel_bytes):
    """Lê o formulário vertical e regista a posição de todos os campos (rótulos)."""
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    ws = wb.active
    
    mapa_campos = {}
    
    # Percorre as linhas do Excel (ex: até à linha 50) e as colunas A e B
    for r in range(1, 50):
        for c in range(1, 3): 
            val = ws.cell(row=r, column=c).value
            if val and isinstance(val, str) and len(val.strip()) > 2:
                # Evita apanhar títulos de secção que costumam ter números (ex: "1. Background")
                if not re.match(r"^\d\.", val.strip()): 
                    mapa_campos[val.strip()] = {"row": r, "col": c}
                    
    return mapa_campos

# ============================================================
# 4. AGENTE DE INTELIGÊNCIA ARTIFICIAL
# ============================================================
def extrair_dados_formulario_com_gemini(texto_notas, campos_disponiveis):
    model = genai.GenerativeModel(model_name=MODELO_GEMINI)
    
    lista_campos = list(campos_disponiveis.keys())

    prompt_sistema = f"""
    Tu és um Assistente de RH encarregue de preencher formulários de avaliação de entrevistas.
    Abaixo tens as notas soltas tiradas durante uma entrevista.
    
    O formulário Excel que tens de preencher tem exatamente os seguintes campos:
    {lista_campos}
    
    O teu objetivo:
    - Extrai a informação do texto e mapeia para os campos corretos.
    - Se não existir informação nas notas para um determinado campo, deixa o valor vazio "".
    - Resume as informações de forma profissional, mantendo os detalhes técnicos.
    
    MUITO IMPORTANTE:
    Responde APENAS com um ÚNICO objeto JSON (e não uma lista/array). As chaves do JSON têm de ser EXATAMENTE as que listei acima.
    """

    try:
        response = model.generate_content(
            [prompt_sistema, f"Notas da Entrevista:\n\n{texto_notas}"],
            generation_config={"temperature": 0.2},
        )

        texto_resposta = (response.text or "").strip()
        if "```" in texto_resposta:
            texto_resposta = re.sub(r"```json|```", "", texto_resposta).strip()

        return json.loads(texto_resposta)
    except Exception as e:
        st.error(f"Erro na interpretação da IA: {e}")
        return None

# ============================================================
# 5. ESCRITA NO FORMULÁRIO EXCEL
# ============================================================
def preencher_formulario_excel(excel_bytes, dados_json, mapa_campos):
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    ws = wb.active

    campos_preenchidos = 0
    for nome_campo, valor in dados_json.items():
        if nome_campo in mapa_campos and valor:
            linha = mapa_campos[nome_campo]["row"]
            coluna_rotulo = mapa_campos[nome_campo]["col"]
            
            # Escreve na célula imediatamente à direita do nome do campo!
            # Exemplo: Se o rótulo está em A5, escrevemos a resposta em B5.
            coluna_destino = coluna_rotulo + 1 
            
            ws.cell(row=linha, column=coluna_destino).value = valor
            campos_preenchidos += 1

    output_final = io.BytesIO()
    wb.save(output_final)
    output_final.seek(0)
    return output_final, campos_preenchidos

# ============================================================
# 6. INTERFACE DO UTILIZADOR
# ============================================================
col1, col2 = st.columns([1, 1])

with col1:
    st.markdown("### 📄 1. Notas da Entrevista")
    ficheiros_carregados = st.file_uploader("Carrega as notas (TXT/DOCX)", type=["txt", "docx"])

    texto_acumulado = ""
    if ficheiros_carregados:
        if ficheiros_carregados.name.endswith(".txt"):
            texto_acumulado = ler_ficheiro_txt(ficheiros_carregados)
        elif ficheiros_carregados.name.endswith(".docx"):
            texto_acumulado = ler_ficheiro_docx(ficheiros_carregados)

    texto_colado = st.text_area("Ou escreve/cola aqui as notas do candidato:", height=200, value=texto_acumulado)

with col2:
    st.markdown("### 📂 2. Formulário de Avaliação (Excel)")
    excel_modelo = st.file_uploader("Carrega o 'Interview evaluation form.xlsx'", type=["xlsx"])
    
    mapa_campos = {}
    if excel_modelo:
        excel_bytes_lido = excel_modelo.read()
        mapa_campos = ler_estrutura_formulario(excel_bytes_lido)
        
        if mapa_campos:
            st.success(f"✅ Formulário reconhecido! Encontrados {len(mapa_campos)} campos.")
        else:
            st.error("❌ Não foi possível ler a estrutura do formulário.")

st.divider()

if st.button("🤖 Analisar Notas e Preencher Formulário", use_container_width=True, type="primary"):
    if not texto_colado:
        st.warning("Insere as notas da entrevista primeiro.")
    elif not excel_modelo or not mapa_campos:
        st.error("Carrega o formulário de Excel primeiro.")
    else:
        with st.spinner("A cruzar as notas do candidato com os campos do formulário..."):
            dados_extraidos = extrair_dados_formulario_com_gemini(texto_colado, mapa_campos)

        if dados_extraidos:
            st.session_state.dados_extraidos = dados_extraidos
            st.session_state.excel_pronto = None 
            st.success("✅ O rascunho do formulário foi gerado! Revê abaixo.")

if st.session_state.dados_extraidos:
    st.markdown("### 👀 3. Revisão do Formulário")
    st.caption("Podes editar os dados extraídos antes de os embutir no Excel original.")
    
    # Mostra os dados em forma de tabela vertical para ser fácil de rever
    df_preview = pd.DataFrame(list(st.session_state.dados_extraidos.items()), columns=["Campo do Formulário", "Valor a Inserir"])
    df_editado = st.data_editor(df_preview, use_container_width=True, hide_index=True)

    if st.button("📥 Embutir Dados e Gerar Novo Excel", use_container_width=True, type="primary"):
        # Converte a tabela editada de volta para o formato Dicionário
        lista_dados_final = dict(zip(df_editado["Campo do Formulário"], df_editado["Valor a Inserir"]))
        
        try:
            excel_modelo.seek(0)
            output_final, num_inseridos = preencher_formulario_excel(
                excel_modelo.read(), lista_dados_final, mapa_campos
            )

            st.session_state.excel_pronto = output_final.getvalue()
            st.session_state.nome_ficheiro_saida = f"Candidato_Avaliado_{datetime.now().strftime('%d-%m-%Y_%H%M')}.xlsx"
            st.success(f"✨ Sucesso! {num_inseridos} campos do formulário foram preenchidos.")
        except Exception as ex:
            st.error(f"Erro a gerar ficheiro: {ex}")

if st.session_state.excel_pronto:
    st.download_button(
        label=f"📥 Descarregar Formulário Preenchido ('{st.session_state.nome_ficheiro_saida}')",
        data=st.session_state.excel_pronto,
        file_name=st.session_state.nome_ficheiro_saida,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
