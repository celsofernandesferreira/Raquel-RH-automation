import streamlit as st
import google.generativeai as genai
import json
import re
import io
import docx
import logging
from datetime import datetime
import openpyxl  # Biblioteca crucial para abrir e editar o Excel existente

# 1. CONFIGURAÇÃO DE LOGS
logging.basicConfig(
    filename="automacao_rh_raquel.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    encoding="utf-8"
)

st.set_page_config(page_title="Automação de Relatórios RH", page_icon="📊", layout="wide")
st.title("📊 Sistema de Automação de Relatórios — RH")
st.subheader("Extração de Notas para Excel Modelo Predefinido")

# 2. INICIALIZAÇÃO DA API DO GEMINI
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
except Exception:
    st.error("Erro: Chave API em falta nos Secrets do Streamlit.")
    st.stop()

# 3. LEITURA DE TEXTOS (Word e TXT)
def ler_ficheiro_txt(uploaded_file):
    return uploaded_file.read().decode("utf-8")

def ler_ficheiro_docx(uploaded_file):
    doc = docx.Document(io.BytesIO(uploaded_file.read()))
    return "\n".join([para.text for para in doc.paragraphs])

# 4. AGENTE DE INTELIGÊNCIA ARTIFICIAL (Com Esquema Estruturado Universal)
def extrair_dados_com_gemini(texto_notas, mapeamento_campos):
    colunas_alvo = [c.strip() for c in mapeamento_campos.split(",")]
    propriedades_json = {campo: {"type": "STRING"} for campo in colunas_alvo}
    
    esquema_resposta = {
        "type": "ARRAY",
        "items": {
            "type": "OBJECT",
            "properties": propiedades_json,
            "required": colunas_alvo
        }
    }

    # Correção do Erro 404: Tenta usar a nomenclatura universal estável 'gemini-pro' 
    # ou garante retrocompatibilidade se a biblioteca for mais antiga.
    try:
        model = genai.GenerativeModel(
            model_name="gemini-pro", 
            generation_config={
                "response_mime_type": "application/json",
                "response_schema": esquema_resposta
            }
        )
    except Exception:
        # Fallback caso a versão do pacote exija a string com o prefixo herdado
        model = genai.GenerativeModel(model_name="gemini-pro")

    prompt_sistema = f"""
    Tu és um Assistente de RH Avançado especialista em extração de dados estruturados. 
    O teu objetivo exclusivo é ler as notas fornecidas e extrair as informações relevantes para preencher a tabela.
    
    Regras estritas:
    1. Retorna APENAS um array JSON válido de objetos com os campos solicitados.
    2. Se a informação para algum dos campos não existir nas notas, deixa o valor como string vazia "".
    3. Nunca inventes dados.
    """
    
    try:
        response = model.generate_content([prompt_sistema, f"Notas para analisar:\n\n{texto_notas}"])
        # Limpeza robusta de markdown caso o modelo não suporte o parâmetro response_schema de forma nativa
        json_clean = response.text.strip()
        if "```json" in json_clean:
            json_clean = re.sub(r"```json|```", "", json_clean).strip()
        return json.loads(json_clean)
    except Exception as e:
        logging.error(f"Erro na extração do Agente: {e}")
        st.sidebar.error(f"Detalhe do erro do Gemini: {e}")
        return None

# 5. INTERFACE DO UTILIZADOR
col1, col2 = st.columns([1, 1])

with col1:
    st.markdown("### 📄 1. Documentos Base (Word/Notepad)")
    ficheiros_carregados = st.file_uploader("Carrega as notas de RH", type=["txt", "docx"], accept_multiple_files=True)
    
    texto_acumulado = ""
    if ficheiros_carregados:
        for f in ficheiros_carregados:
            if f.name.endswith(".txt"):
                texto_acumulado += ler_ficheiro_txt(f) + "\n"
            elif f.name.endswith(".docx"):
                texto_acumulado += ler_ficheiro_docx(f) + "\n"
                
    texto_colado = st.text_area("Ou escreve/cola texto adicional aqui:")
    texto_final = (texto_acumulado + "\n" + texto_colado).strip()

with col2:
    st.markdown("### 📂 2. Configuração do Modelo da Empresa")
    excel_modelo = st.file_uploader("Carrega aqui o Excel Modelo/Template da Empresa (.xlsx)", type=["xlsx"])
    
    st.caption("Identificação das Colunas (Escreve os nomes exatamente como aparecem no cabeçalho do Excel original, separados por vírgulas):")
    campos_predefinidos = "Nome, Data, Ocorrência, Detalhes"
    campos_usuario = st.text_input("Colunas do Excel original:", value=campos_predefinidos)

st.divider()

# 6. PROCESSAMENTO E INJEÇÃO NO TEMPLATE
if st.button("🤖 Analisar Notas e Preencher Excel da Empresa", use_container_width=True, type="primary"):
    if not texto_final:
        st.warning("Introduz ou carrega notas primeiro.")
    elif not excel_modelo:
        st.error("Precisas de carregar o Excel Modelo da empresa para podermos preenchê-lo.")
    else:
        with st.spinner("O Agente está a ler o documento e a mapear as células no teu Modelo..."):
            
            lista_dados = extrair_dados_com_gemini(texto_final, campos_usuario)
            
            if lista_dados:
                try:
                    wb = openpyxl.load_workbook(io.BytesIO(excel_modelo.read()))
                    ws = wb.active
                    
                    colunas_alvo = [c.strip() for c in campos_usuario.split(",")]
                    
                    # 1. Encontrar em que linha estão os cabeçalhos
                    linha_cabecalho = 1
                    mapa_colunas_index = {}
                    
                    for r in range(1, 21):
                        valores_linha = []
                        for c in range(1, ws.max_column + 1):
                            cell = ws.cell(row=r, column=c)
                            # Se for uma célula fundida (MergedCell), ignoramos o erro de leitura
                            val = str(cell.value).strip() if cell.value is not None else ""
                            valores_linha.append(val)
                        
                        if any(campo in valores_linha for campo in colunas_alvo):
                            linha_cabecalho = r
                            for c in range(1, ws.max_column + 1):
                                cell = ws.cell(row=r, column=c)
                                nome_celula = str(cell.value).strip() if cell.value is not None else ""
                                if nome_celula in colunas_alvo:
                                    mapa_colunas_index[nome_celula] = c
                            break
                    
                    if not mapa_colunas_index:
                        for idx, nome_col in enumerate(colunas_alvo, start=1):
                            mapa_colunas_index[nome_col] = idx
                    
                    # 2. Descobrir a próxima linha vazia (Prevenindo erros de MergedCells)
                    proxima_linha = linha_cabecalho + 1
                    while True:
                        linha_vazia = True
                        for c in range(1, ws.max_column + 1):
                            cell = ws.cell(row=proxima_linha, column=c)
                            # Verifica se o tipo do objeto é MergedCell ou tem valor
                            if type(cell).__name__ == "MergedCell" or cell.value is not None:
                                linha_vazia = False
                                break
                        if linha_vazia:
                            break
                        proxima_linha += 1
                    
                    # 3. Inserir os dados protegendo contra escrita em células bloqueadas/fundidas
                    linhas_inseridas = 0
                    for registo in lista_dados:
                        # Certifica que saltamos qualquer MergedCell residual para evitar o erro "read-only"
                        while any(type(ws.cell(row=proxima_linha, column=num_col)).name == "MergedCell" for num_col in mapa_colunas_index.values()):
                            proxima_linha += 1

                        for nome_campo, num_coluna in mapa_colunas_index.items():
                            valor_final = registo.get(nome_campo, "")
                            cell = ws.cell(row=proxima_linha, column=num_coluna)
                            
                            # Só escreve se NÃO for uma célula fundida secundária
                            if type(cell).__name__ != "MergedCell":
                                cell.value = valor_final
                                
                        proxima_linha += 1
                        linhas_inseridas += 1
                    
                    output_final = io.BytesIO()
                    wb.save(output_final)
                    output_final.seek(0)
                    
                    st.success(f"✨ Sucesso! Adicionámos {linhas_inseridas} novas linhas ao documento original sem alterar a estrutura.")
                    
                    st.download_button(
                        label="📥 Descarregar Excel Preenchido",
                        data=output_final,
                        file_name=f"Relatorio_Final_{datetime.now().strftime('%d-%m-%Y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                except Exception as ex:
                    st.error(f"Erro ao manipular o ficheiro Excel: {ex}")
                    logging.error(f"Erro na manipulação de openpyxl: {ex}")
            else:
                st.error("Não foi possível extrair dados válidos. Verifica a barra lateral para ver o detalhe do erro.")
                
