import streamlit as st
import pandas as pd
import numpy as np
from langchain_openai import ChatOpenAI
from langchain_experimental.agents.agent_toolkits import create_pandas_dataframe_agent
import os
import matplotlib.pyplot as plt
import httpx
import io
import msoffcrypto

# Configura o backend do matplotlib para não travar o Streamlit
plt.switch_backend('Agg')

# =================================================================
# 1. CONFIGURAÇÕES GERAIS
# =================================================================
OPENAI_API_KEY = st.secrets["OPENAI_API_KEY"]

st.set_page_config(
    page_title="Unipar | Central Agente Manuntenção (Piloto)",
    page_icon="🏭",
    layout="wide"
)

# --- ESTILIZAÇÃO CSS ---
st.markdown(f"""
    <style>
    /* 1. Header Principal com Gradiente */
    .header-container {{
        background: linear-gradient(135deg, #008D36 0%, #005C23 100%);
        padding: 1.5rem;
        border-radius: 12px;
        color: white !important;
        text-align: center;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 10px rgba(0,0,0,0.15);
    }}
    .header-container h3 {{
        color: #FFFFFF !important;
        font-weight: 700;
        margin: 0;
    }}
    .header-container p {{
        color: #F0F0F0 !important;
        margin-top: 5px;
        font-size: 1rem;
    }}

    /* 2. Ajuste dos Balões de Chat */
    [data-testid="stChatMessage"] {{
        background-color: #F8F9FA;
        border: 1px solid #D0D7DE;
        border-radius: 12px;
        padding: 1.2rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        color: #1F2328;
        margin-bottom: 10px;
    }}
    
    [data-testid="stChatMessage"] .stMarkdown {{
        color: #1F2328 !important;
    }}

    /* Ícones dos avatares */
    .st-emotion-cache-1p1m4t5 {{
        background-color: #008D36 !important;
        color: white !important;
    }}
    </style>
    """, unsafe_allow_html=True)

# =================================================================
# 2. MOTOR DE DADOS
# =================================================================
@st.cache_data(show_spinner=False)
def carregar_sistema_pcm(filepath, password):
    """
    Carrega um arquivo Excel protegido por senha usando msoffcrypto.
    """
    if not os.path.exists(filepath):
        return None, f"Arquivo '{filepath}' não encontrado no diretório."

    # Abre o arquivo criptografado
    with open(filepath, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password=password) # Define a senha

        # Descriptografa para memória (buffer)
        decrypted_file = io.BytesIO()
        office_file.decrypt(decrypted_file)

    # Lê o conteúdo descriptografado com pandas
    # Se sua planilha tiver nome específico, adicione sheet_name='Planilha1' aqui
    df = pd.read_excel(decrypted_file) 
    return df, None

# --- CONFIGURAÇÃO DO ARQUIVO LOCAL ---
ARQUIVO_LOCAL = "dados1.xlsx"
SENHA_ARQUIVO = st.secrets["senha"]
df, error_msg = carregar_sistema_pcm(ARQUIVO_LOCAL, SENHA_ARQUIVO)

# =================================================================
# 3. DEFINIÇÃO DA IA (CÉREBRO)
# =================================================================


dicionario_dados = st.secrets["dicionario_dados"]


# Adicionei instruções explícitas sobre o nome do arquivo
instrucoes_pcm = f"""
Você é um Especialista Sênior em Planejamento e Controle de Manutenção (PCM) da Unipar.
Sua missão é extrair indicadores da base carregada no dataframe `df` usando Python (Pandas).

### 1. ENTENDIMENTO DA BASE (CRÍTICO)
- **Dicionário:**
{dicionario_dados}
- **Granularidade:** Considere que cada linha do dataframe representa uma operação ou uma ordem de serviço individual.
- **Regra de Contagem (OURO):** A menos que o usuário peça explicitamente por "contagem de ordens únicas/distintas", use sempre a contagem total de linhas (`len(df)` ou `.count()`). Não use `.nunique()` para volumes gerais, pois isso esconde o volume real de trabalho.

### 2. REGRAS DE CÁLCULO (PYTHON PANDAS)
Traduza as perguntas do usuário para o código abaixo. Não altere a lógica.

- **TOTAL DE ORDENS (Volume Geral):**
  `df_filtrado = df[df['Ordem'].notna()]`
  Resultado = `len(df_filtrado)`

- **TOTAL DE ORDENS EXECUTADAS:**
  `df_exec = df[df['Status usuário'] == 'EXEC']`
  Resultado = `len(df_exec)`

- **ORDENS VENCIDAS (Backlog Crítico):**
  Critério: Atrasada, não executada e não cancelada.
  `mask_vencida = (df['Status Prazo'] == 'Vencida') & (~df['Status usuário'].isin(['EXEC', 'CANC']))`
  `df_vencidas = df[mask_vencida]`
  Resultado = `len(df_vencidas)`

- **NOTAS EM ABERTO (Carteira de Serviços):**
  Critério: Tem Nota mas não tem Ordem criada.
  `df_notas = df[(df['Nota'].notna()) & (df['Ordem'].isna())]`
  Resultado = `len(df_notas)`

### 3. DIRETRIZES DE ANÁLISE
1. **Analise antes de responder:** Se o usuário perguntar "Quantas ordens vencidas?", rode o código Python, veja o número e responda com o número exato.
2. **Datas:** Para perguntas de tempo ("Evolução", "Por mês"), use a coluna `Date`.
3. **Agrupamentos:** Se perguntarem "Quais áreas tem mais atraso?", faça um `value_counts()` na coluna `ÁREA` filtrando pelas vencidas.
4. **Restrição:** NÃO invente dados. Se a coluna não existe ou está vazia, informe.
5. **Formatação:** Apresente números grandes com separador de milhar (ex: 1.234) e percentuais com 1 casa decimal.


### REGRAS PARA GRÁFICOS:
1. Se o usuário pedir um gráfico, use matplotlib.
2. **OBRIGATÓRIO:** Salve o gráfico SEMPRE com o nome exato: `temp_chart.png` usando `plt.savefig('temp_chart.png')`.
3. Sempre limpe a figura com `plt.clf()` ou `plt.close()` após salvar para evitar sobreposição.
4. Informe na resposta: "Gráfico gerado com sucesso."


Agora, aguarde a pergunta do usuário e execute o código Python necessário na variável `df`.
"""
# =================================================================
# 3. BARRA LATERAL (SIMPLIFICADA)
# =================================================================
with st.sidebar:
    st.markdown("### ⚙️ Painel Agente de Manuntenção")
    
    st.warning("🚧 **PROJETO PILOTO**\n\nAmbiente para coleta de feedbacks. Respostas geradas por IA.")
    st.divider()

    st.info("🔒 Status do Sistema: **Online**")
    
    if error_msg:
        st.error("⚠️ Atenção: Falha na conexão com a base de dados.")
    
    st.markdown("---")
    st.caption("© 2026 Unipar Data Analytics")
    
   
# =================================================================
# 4. INTERFACE DE CHAT
# =================================================================

st.markdown("""
<div class="header-container">
    <h3>Agente de Inteligência PCM</h3>
    <p>Assistente Virtual para Análise de Manutenção e Backlog</p>
</div>
""", unsafe_allow_html=True)

if "messages_pcm" not in st.session_state:
    st.session_state.messages_pcm = [
        {"role": "assistant", "content": "Olá! Sou seu assistente de PCM. Como posso ajudar?"}
    ]

# Exibe Histórico
for msg in st.session_state.messages_pcm:
    with st.chat_message(msg["role"]):
        st.write(msg["content"])
        if "image" in msg and msg["image"] is not None:
            st.image(msg["image"])

# Input do Usuário
if prompt_pcm := st.chat_input("Ex: Gere um gráfico de barras com as Ordens Vencidas por Área."):
    
    if df is None:
        st.error(f"Erro na base de dados: {error_msg}")
    else:
        st.session_state.messages_pcm.append({"role": "user", "content": prompt_pcm})
        st.chat_message("user").write(prompt_pcm)

        with st.spinner("Analisando indicadores..."):
            # Limpa gráfico anterior se existir para não confundir
            if os.path.exists("temp_chart.png"):
                os.remove("temp_chart.png")

            try:                
                http_client_inseguro = httpx.Client(verify=False)
                llm = ChatOpenAI(
                    model="gpt-4o",
                    api_key=OPENAI_API_KEY,                     
                    temperature=0,
                    http_client=http_client_inseguro
                )
                
                agente_manutencao = create_pandas_dataframe_agent(
                    llm,
                    df,
                    prefix=instrucoes_pcm,
                    verbose=True,
                    allow_dangerous_code=True,
                    agent_type="openai-functions"
                )

                resposta = agente_manutencao.run(prompt_pcm)
                
                # Verifica se a IA gerou imagem
                imagem_para_historico = None
                if os.path.exists("temp_chart.png"):
                    with open("temp_chart.png", "rb") as f:
                        imagem_para_historico = f.read()
                
                # Adiciona ao histórico e exibe
                msg_aux = {"role": "assistant", "content": resposta, "image": imagem_para_historico}
                st.session_state.messages_pcm.append(msg_aux)
                
                with st.chat_message("assistant"):
                    st.write(resposta)
                    if imagem_para_historico:
                        st.image(imagem_para_historico)
                
                # Limpa o arquivo temporário após carregar na memória do session_state
                if os.path.exists("temp_chart.png"):
                    os.remove("temp_chart.png")

            except Exception as e:
                st.error(f"Erro: {e}")
