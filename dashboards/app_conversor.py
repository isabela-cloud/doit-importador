"""
Interface Streamlit - Conversor de Dados para Implantação
Upload de arquivos de clientes → conversão para padrão DOit
"""

import streamlit as st
import pandas as pd
import os
import sys
import io

def _get_engine(filename):
    """Retorna o engine correto para leitura de Excel baseado na extensão."""
    if filename.endswith(".xls"):
        return "xlrd"
    return "openpyxl"

# Adicionar pasta scripts ao path
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'scripts'))

from conversor import (
    converter_arquivo,
    obter_preview_mapeamento,
    carregar_modelo,
    MAPEAMENTOS,
    TIPOS_RECEITA_DESPESA,
    _encontrar_coluna,
    _mapear_automatico,
    _aplicar_formatacoes,
)

# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================
st.set_page_config(
    page_title="Conversor de Dados - Implantação",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main .block-container {padding-top: 1rem;}
    .stSuccess {border-radius: 10px;}
</style>
""", unsafe_allow_html=True)

# ============================================================
# SIDEBAR
# ============================================================
LOGO_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), '..', 'implantacao-dashboard', 'templates', 'DOit logo.png')
if os.path.exists(LOGO_PATH):
    st.sidebar.image(LOGO_PATH, width=100)

st.sidebar.title("🔄 Conversor de Dados")

# Seleção do tipo de dado
tipo_opcoes = {
    'Contatos / Pessoas': 'contatos',
    'Projetos': 'projetos',
    'Financeiro': 'financeiro',
    'Horas Trabalhadas': 'horas',
    'Usuários': 'usuarios',
    'Produtos': 'produtos',
    'Vendas': 'vendas',
}
tipo_label = st.sidebar.selectbox("📋 Tipo de dado", list(tipo_opcoes.keys()))
tipo = tipo_opcoes[tipo_label]

# Seleção da origem
origem_opcoes = {
    'DOit Coleta (Planilha Padrão)': 'doit_coleta',
    'Conta Azul': 'conta_azul',
    'Navis': 'navis',
    'ClickUp': 'clickup',
    'Sienge': 'sienge',
    'Trello': 'trello',
    'Excel Desestruturado': 'excel_desestruturado',
    'Financeiro Horizontal (BPO)': 'financeiro_horizontal',
    'Excel Manual / Outro': 'excel_manual',
}
origem_label = st.sidebar.selectbox("🔗 Sistema de origem", list(origem_opcoes.keys()))
origem = origem_opcoes[origem_label]

st.sidebar.markdown("---")
st.sidebar.subheader("⚙️ Configurações")

# Estilo de caixa
estilo_caixa_opcao = st.sidebar.radio(
    "🔤 Estilo de texto",
    ['Primeira Maiúscula', 'TUDO MAIÚSCULO', 'Original (não alterar)'],
    index=0,
    help="Define como os campos de texto serão formatados"
)
if estilo_caixa_opcao == 'TUDO MAIÚSCULO':
    estilo_caixa = 'MAIÚSCULA'
elif estilo_caixa_opcao == 'Original (não alterar)':
    estilo_caixa = 'ORIGINAL'
else:
    estilo_caixa = 'Primeira Maiúscula'

# ID inicial por tipo
st.sidebar.markdown("---")
st.sidebar.subheader("🔢 ID inicial por tipo")
id_inicial_contatos = st.sidebar.number_input("Contatos", min_value=1, value=15, step=1, key='id_contatos')
id_inicial_projetos = st.sidebar.number_input("Projetos", min_value=1, value=2, step=1, key='id_projetos')
id_inicial_financeiro = st.sidebar.number_input("Financeiro", min_value=1, value=47, step=1, key='id_financeiro')
id_inicial_horas = st.sidebar.number_input("Horas", min_value=1, value=1, step=1, key='id_horas')
id_inicial_usuarios = st.sidebar.number_input("Usuários", min_value=1, value=1, step=1, key='id_usuarios')
id_inicial_produtos = st.sidebar.number_input("Produtos", min_value=1, value=1, step=1, key='id_produtos')
id_inicial_vendas = st.sidebar.number_input("Vendas", min_value=1, value=1, step=1, key='id_vendas')

# Selecionar o ID correto para o tipo atual
id_inicial_map = {
    'contatos': id_inicial_contatos,
    'projetos': id_inicial_projetos,
    'financeiro': id_inicial_financeiro,
    'horas': id_inicial_horas,
    'usuarios': id_inicial_usuarios,
    'produtos': id_inicial_produtos,
    'vendas': id_inicial_vendas,
}
id_inicial = id_inicial_map[tipo]

# Upload de referências para vínculo de IDs
cadastro_ref_file = None
projetos_ref_file = None

if tipo in ['projetos', 'financeiro']:
    st.sidebar.subheader("🔗 Vincular IDs")
    cadastro_ref_file = st.sidebar.file_uploader(
        "Cadastro convertido (para ID)",
        type=['xlsx'],
        key='cadastro_ref',
        help="Upload do cadastro já convertido para vincular ID DO CADASTRO"
    )

if tipo == 'financeiro':
    projetos_ref_file = st.sidebar.file_uploader(
        "Projetos convertido (para ID)",
        type=['xlsx'],
        key='projetos_ref',
        help="Upload dos projetos já convertidos para vincular ID PROJETO"
    )

st.sidebar.markdown("---")
st.sidebar.markdown("""
**Como usar:**
1. Selecione o tipo e origem
2. Configure caixa e IDs iniciais
3. Faça upload do arquivo
4. Revise o mapeamento
5. Baixe o arquivo convertido
""")

# ============================================================
# ÁREA PRINCIPAL
# ============================================================
st.title("🔄 Conversor de Dados para Implantação")
st.markdown("Converta arquivos de clientes para o padrão DOit de forma rápida e segura.")
st.markdown("---")

# Upload do arquivo
uploaded_file = st.file_uploader(
    "📂 Faça upload do arquivo do cliente",
    type=['xlsx', 'xls', 'csv'],
    help="Aceita arquivos Excel (.xlsx, .xls) e CSV (.csv)"
)

if uploaded_file is not None:
    # Ler arquivo
    try:
        if uploaded_file.name.endswith('.csv'):
            df_entrada = pd.read_csv(uploaded_file)
        else:
            # Se for Financeiro Horizontal (BPO), usar parser especial
            if origem == 'financeiro_horizontal':
                from parser_bpo import parse_financeiro_horizontal
                
                xls = pd.ExcelFile(uploaded_file, engine=_get_engine(uploaded_file.name))
                abas = xls.sheet_names
                aba_escolhida = st.selectbox("📄 Selecione a aba com o financeiro", abas, key='aba_bpo')
                
                ano_bpo = st.number_input("Ano de referência", min_value=2020, max_value=2030, 
                                          value=2026, step=1, key='ano_bpo')
                
                try:
                    # Salvar temporariamente para o parser ler
                    import tempfile
                    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                        uploaded_file.seek(0)
                        tmp.write(uploaded_file.read())
                        tmp_path = tmp.name
                    
                    df_entrada = parse_financeiro_horizontal(tmp_path, sheet_name=aba_escolhida, ano=ano_bpo)
                    
                    import os as _os
                    _os.unlink(tmp_path)
                    
                    st.info(f"📊 {len(df_entrada)} lançamentos extraídos do BPO")
                except Exception as e:
                    st.error(f"❌ Erro ao processar financeiro horizontal: {e}")
                    st.stop()
            
            # Se for Navis, usar parser especial
            elif origem == 'navis':
                from parser_navis import parse_navis, _detectar_tipo_relatorio
                import tempfile
                
                xls = pd.ExcelFile(uploaded_file, engine=_get_engine(uploaded_file.name))
                abas = xls.sheet_names
                
                if len(abas) > 1:
                    aba_escolhida = st.selectbox("📄 Selecione a aba", abas, key='aba_navis')
                else:
                    aba_escolhida = abas[0]
                
                # Detectar tipo de relatório automaticamente
                uploaded_file.seek(0)
                df_detect = pd.read_excel(uploaded_file, engine=_get_engine(uploaded_file.name), sheet_name=aba_escolhida, header=None, nrows=15)
                tipo_rel_detectado = _detectar_tipo_relatorio(df_detect)
                
                # Permitir override manual
                opcoes_tipo_navis = {
                    'financeiro_cc': 'Movimentos de Conta Corrente',
                    'financeiro_previsao': 'Contas a Pagar/Receber (Previsão)',
                    'projetos': 'Consulta Projetos - Cadastro',
                    'clientes': 'Clientes',
                    'fornecedores': 'Fornecedores',
                    'contatos': 'Contatos',
                    'horas': 'Aplicação de Horas',
                }
                
                tipo_rel_label = st.selectbox(
                    "📋 Tipo de relatório Navis (detectado automaticamente)",
                    list(opcoes_tipo_navis.values()),
                    index=list(opcoes_tipo_navis.keys()).index(tipo_rel_detectado) if tipo_rel_detectado in opcoes_tipo_navis else 0,
                    key='tipo_navis'
                )
                tipo_rel = [k for k, v in opcoes_tipo_navis.items() if v == tipo_rel_label][0]
                
                # Validar compatibilidade tipo de dado × tipo de relatório
                if tipo == 'contatos' and tipo_rel not in ('clientes', 'fornecedores', 'contatos'):
                    st.warning("⚠️ Para converter cadastros do Navis, use um relatório de Clientes, Fornecedores ou Contatos.")
                elif tipo == 'horas' and tipo_rel != 'horas':
                    st.warning("⚠️ Para converter horas do Navis, use o relatório de Aplicação de Horas.")
                
                try:
                    # Salvar temporariamente para o parser ler
                    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                        uploaded_file.seek(0)
                        tmp.write(uploaded_file.read())
                        tmp_path = tmp.name
                    
                    df_entrada = parse_navis(tmp_path, sheet_name=aba_escolhida, tipo_relatorio=tipo_rel)
                    
                    import os as _os
                    _os.unlink(tmp_path)
                    
                    st.info(f"📊 **{len(df_entrada)} registros** extraídos do relatório Navis ({opcoes_tipo_navis[tipo_rel]})")
                    
                    # Se o tipo de dado é financeiro mas o relatório é de projetos, avisar
                    if tipo == 'financeiro' and tipo_rel in ('projetos', 'clientes', 'fornecedores', 'contatos', 'horas'):
                        st.warning("⚠️ Você selecionou tipo 'Financeiro' mas o relatório não é financeiro. Ajuste o tipo de dado na sidebar.")
                    elif tipo == 'projetos' and tipo_rel not in ('projetos',):
                        st.warning("⚠️ Você selecionou tipo 'Projetos' mas o relatório não é de projetos. Ajuste o tipo de dado na sidebar.")
                    elif tipo == 'contatos' and tipo_rel not in ('clientes', 'fornecedores', 'contatos'):
                        st.warning("⚠️ Você selecionou tipo 'Contatos' mas o relatório não é de cadastro. Ajuste o tipo de dado na sidebar.")
                    elif tipo == 'horas' and tipo_rel != 'horas':
                        st.warning("⚠️ Você selecionou tipo 'Horas' mas o relatório não é de horas. Ajuste o tipo de dado na sidebar.")
                    
                except Exception as e:
                    st.error(f"❌ Erro ao processar relatório Navis: {e}")
                    st.stop()
            
            # Se for Conta Azul, selecionar aba correta e tratar cabeçalho
            elif origem == 'conta_azul':
                xls = pd.ExcelFile(uploaded_file, engine=_get_engine(uploaded_file.name))
                abas = xls.sheet_names
                
                # Selecionar aba automaticamente com base no tipo
                if tipo == 'financeiro':
                    aba_alvo = next((a for a in abas if 'extrato' in a.lower() or 'financeiro' in a.lower()), None)
                elif tipo == 'contatos':
                    aba_alvo = next((a for a in abas if 'cadastro' in a.lower() or 'cliente' in a.lower()), None)
                else:
                    aba_alvo = None
                
                if aba_alvo is None:
                    aba_alvo = st.selectbox("📄 Selecione a aba com os dados", abas, key='aba_conta_azul')
                else:
                    st.info(f"📄 Aba detectada automaticamente: **{aba_alvo}**")
                
                # Ler sem header para detectar cabeçalho real
                df_raw = pd.read_excel(uploaded_file, engine=_get_engine(uploaded_file.name), sheet_name=aba_alvo, header=None)
                
                # Detectar linha de cabeçalho (Conta Azul pode ter linha de índice numérico)
                palavras_chave_ca = ['descrição', 'valor', 'data', 'tipo', 'categoria', 'situação',
                                     'fornecedor', 'cliente', 'nome', 'cnpj', 'cpf', 'email',
                                     'conta bancária', 'forma de pgto', 'vencimento', 'movimento']
                
                melhor_linha = 0
                melhor_score = 0
                for i in range(min(10, len(df_raw))):
                    row_vals = [str(v).lower().strip() for v in df_raw.iloc[i] if pd.notna(v) and str(v).strip()]
                    score = sum(1 for val in row_vals for kw in palavras_chave_ca if kw in val)
                    score += len(row_vals) * 0.1
                    if score > melhor_score:
                        melhor_score = score
                        melhor_linha = i
                
                df_entrada = pd.read_excel(uploaded_file, engine=_get_engine(uploaded_file.name), sheet_name=aba_alvo, header=melhor_linha)
                
                # Limpar colunas Unnamed
                df_entrada = df_entrada.loc[:, df_entrada.columns.notna()]
                df_entrada = df_entrada[[c for c in df_entrada.columns if 'unnamed' not in str(c).lower()]]
                
                # Remover colunas com nome numérico (índice do Conta Azul)
                df_entrada = df_entrada[[c for c in df_entrada.columns if not isinstance(c, (int, float))]]
                
                # Remover linhas completamente vazias
                df_entrada = df_entrada.dropna(how='all').reset_index(drop=True)
                
                # Tratar coluna TIPO: Conta Azul usa "Crédito"/"Débito" → converter para "Receita"/"Despesa"
                if tipo == 'financeiro':
                    col_tipo = next((c for c in df_entrada.columns if 'tipo da opera' in str(c).lower() or c == 'Tipo da operação'), None)
                    if col_tipo:
                        df_entrada[col_tipo] = df_entrada[col_tipo].apply(
                            lambda x: 'Receita' if pd.notna(x) and 'cr' in str(x).lower() 
                            else ('Despesa' if pd.notna(x) and 'déb' in str(x).lower() or (pd.notna(x) and 'deb' in str(x).lower()) else x)
                        )
                    
                    # Tratar CONCILIADO: "Quitado" → "Sim", outros → "Não"
                    col_situacao = next((c for c in df_entrada.columns if 'situa' in str(c).lower()), None)
                    if col_situacao:
                        df_entrada[col_situacao] = df_entrada[col_situacao].apply(
                            lambda x: 'Sim' if pd.notna(x) and 'quit' in str(x).lower() else 'Não'
                        )
                
                st.caption(f"Cabeçalho detectado na linha {melhor_linha + 1} | {len(df_entrada)} registros")
            
            # Se for Excel Desestruturado, mostrar seletor de aba e detectar cabeçalho
            elif origem == 'excel_desestruturado':
                xls = pd.ExcelFile(uploaded_file, engine=_get_engine(uploaded_file.name))
                abas = xls.sheet_names
                
                aba_escolhida = st.selectbox("📄 Selecione a aba com os dados", abas, key='aba_desestruturado')
                
                # Ler a aba sem header para detectar onde começam os dados
                df_raw = pd.read_excel(uploaded_file, engine=_get_engine(uploaded_file.name), sheet_name=aba_escolhida, header=None)
                
                # Detectar linha de cabeçalho: procurar a linha com mais texto não-vazio
                # e que contenha palavras-chave como "nome", "cliente", "projeto", "status", "data"
                palavras_chave = ['nome', 'cliente', 'projeto', 'status', 'data', 'início', 'inicio',
                                  'nível', 'nivel', 'responsável', 'responsavel', 'descrição', 'descricao',
                                  'projetista', 'prioridade', 'categoria']
                
                melhor_linha = 0
                melhor_score = 0
                for i in range(min(20, len(df_raw))):
                    row_vals = [str(v).lower().strip() for v in df_raw.iloc[i] if pd.notna(v) and str(v).strip()]
                    score = sum(1 for val in row_vals for kw in palavras_chave if kw in val)
                    # Bonus por ter muitas colunas preenchidas
                    score += len(row_vals) * 0.1
                    if score > melhor_score:
                        melhor_score = score
                        melhor_linha = i
                
                header_linha = st.number_input(
                    "Linha do cabeçalho (detectada automaticamente)",
                    min_value=0, max_value=len(df_raw)-1, value=melhor_linha,
                    help="Linha onde estão os nomes das colunas (0 = primeira linha)"
                )
                
                # Reler com o header correto
                df_entrada = pd.read_excel(uploaded_file, engine=_get_engine(uploaded_file.name), sheet_name=aba_escolhida, header=int(header_linha))
                
                # Limpar colunas sem nome
                df_entrada = df_entrada.loc[:, df_entrada.columns.notna()]
                df_entrada = df_entrada[[c for c in df_entrada.columns if 'unnamed' not in str(c).lower()]]
                
                # Remover linhas que parecem ser legendas, cabeçalhos repetidos ou vazias
                if not df_entrada.empty:
                    # Remover linhas completamente vazias
                    df_entrada = df_entrada.dropna(how='all').reset_index(drop=True)
                    
                    # Remover linhas que são legendas (contêm "Nível X de Prioridade", "Legenda", "*")
                    if len(df_entrada.columns) > 0:
                        todas_colunas_str = df_entrada.astype(str).apply(lambda row: ' '.join(row), axis=1).str.lower()
                        mask_legenda = (
                            todas_colunas_str.str.contains('nível.*prioridade|legenda|status.*projetista.*nível', regex=True, na=False) |
                            todas_colunas_str.str.match(r'^\s*\*?\s*$', na=False)
                        )
                        df_entrada = df_entrada[~mask_legenda].reset_index(drop=True)
                    
                    # Remover linhas onde a maioria das colunas está vazia (>80%)
                    threshold = len(df_entrada.columns) * 0.8
                    df_entrada = df_entrada[df_entrada.isna().sum(axis=1) < threshold].reset_index(drop=True)
            
            # Se for DOit Coleta, ler a aba correta
            elif origem == 'doit_coleta':
                abas_por_tipo = {
                    'contatos': 'MODELO DE CONTATOS ',
                    'projetos': 'MODELO DE PROJETOS',
                    'financeiro': 'MODELO FINANCEIRO ',
                    'horas': None,
                }
                aba_alvo = abas_por_tipo.get(tipo)
                
                # Verificar abas disponíveis
                xls = pd.ExcelFile(uploaded_file, engine=_get_engine(uploaded_file.name))
                abas_disponiveis = xls.sheet_names
                
                if aba_alvo and aba_alvo in abas_disponiveis:
                    df_entrada = pd.read_excel(uploaded_file, engine=_get_engine(uploaded_file.name), sheet_name=aba_alvo)
                elif aba_alvo:
                    # Tentar match parcial (caso tenha espaço extra ou diferença)
                    aba_encontrada = None
                    for aba in abas_disponiveis:
                        if tipo == 'contatos' and 'contato' in aba.lower():
                            aba_encontrada = aba
                            break
                        elif tipo == 'projetos' and 'projeto' in aba.lower():
                            aba_encontrada = aba
                            break
                        elif tipo == 'financeiro' and 'financeiro' in aba.lower():
                            aba_encontrada = aba
                            break
                        elif tipo == 'horas' and ('hora' in aba.lower() or 'atividade' in aba.lower()):
                            aba_encontrada = aba
                            break
                    
                    if aba_encontrada:
                        df_entrada = pd.read_excel(uploaded_file, engine=_get_engine(uploaded_file.name), sheet_name=aba_encontrada)
                    else:
                        st.warning(f"⚠️ Aba para '{tipo_label}' não encontrada. Abas disponíveis: {abas_disponiveis}")
                        df_entrada = pd.read_excel(uploaded_file, engine=_get_engine(uploaded_file.name), sheet_name=0)
                else:
                    df_entrada = pd.read_excel(uploaded_file, engine=_get_engine(uploaded_file.name), sheet_name=0)
                
                # Remover linhas de DESCRIÇÃO e EXEMPLO (verificar na primeira coluna original)
                if not df_entrada.empty and len(df_entrada.columns) > 0:
                    primeira_col = df_entrada.iloc[:, 0].astype(str).str.strip().str.upper()
                    mask_remover = (
                        primeira_col.str.startswith('DESCRIÇÃO') |
                        primeira_col.str.startswith('EXEMPLO') |
                        primeira_col.str.contains('EXEMPLO', na=False)
                    )
                    df_entrada = df_entrada[~mask_remover].reset_index(drop=True)
                
                # Detectar possíveis registros de exemplo do cliente
                # (dados fictícios que o cliente não apagou da planilha padrão)
                exemplos_conhecidos = [
                    '123.456.708-55', '12.345.678/0001-99', '123.456.789-00',
                    'renato pereira júnior', 'marcenaria pinho',
                    'casa guarujá', 'renatopereira123@gmail.com',
                    'pinhomadeiras@hotmail.com.br',
                ]
                linhas_exemplo = []
                for idx, row in df_entrada.iterrows():
                    row_str = ' '.join(str(v).lower().strip() for v in row if pd.notna(v))
                    for exemplo in exemplos_conhecidos:
                        if exemplo.lower() in row_str:
                            linhas_exemplo.append(idx)
                            break
                
                if linhas_exemplo:
                    st.warning(
                        f"⚠️ **{len(linhas_exemplo)} registro(s) parecem ser exemplos da planilha padrão** "
                        f"(CPF 123.456..., Renato Pereira, etc). Linhas: {[i+1 for i in linhas_exemplo]}"
                    )
                    remover_exemplos = st.checkbox("Remover esses registros de exemplo", value=True)
                    if remover_exemplos:
                        df_entrada = df_entrada.drop(linhas_exemplo).reset_index(drop=True)
                
                # Remover primeira coluna se for None/vazia/índice
                if df_entrada.columns[0] is None or str(df_entrada.columns[0]).strip() == '' or 'unnamed' in str(df_entrada.columns[0]).lower():
                    df_entrada = df_entrada.iloc[:, 1:]
                
                # Remover linhas completamente vazias
                df_entrada = df_entrada.dropna(how='all').reset_index(drop=True)
            
            # Sienge e ClickUp: detectar cabeçalho automaticamente (pular metadados no topo)
            elif origem in ('sienge', 'clickup'):
                df_raw = pd.read_excel(uploaded_file, engine=_get_engine(uploaded_file.name), header=None)
                
                # Procurar linha com mais colunas preenchidas e palavras-chave
                palavras_chave_fin = ['data', 'valor', 'histórico', 'documento', 'débito', 'crédito',
                                      'conta', 'emissão', 'vencto', 'baixa', 'favorecido', 'classificação']
                
                melhor_linha = 0
                melhor_score = 0
                for i in range(min(15, len(df_raw))):
                    row_vals = [str(v).lower().strip() for v in df_raw.iloc[i] if pd.notna(v) and str(v).strip()]
                    score = sum(1 for val in row_vals for kw in palavras_chave_fin if kw in val)
                    score += len(row_vals) * 0.2
                    if score > melhor_score:
                        melhor_score = score
                        melhor_linha = i
                
                if melhor_score > 1:
                    df_entrada = pd.read_excel(uploaded_file, engine=_get_engine(uploaded_file.name), header=melhor_linha)
                    # Limpar colunas Unnamed
                    df_entrada = df_entrada[[c for c in df_entrada.columns if 'unnamed' not in str(c).lower()]]
                    # Remover linhas vazias
                    df_entrada = df_entrada.dropna(how='all').reset_index(drop=True)
                    st.caption(f"Cabeçalho detectado na linha {melhor_linha + 1}")
                else:
                    df_entrada = pd.read_excel(uploaded_file, engine=_get_engine(uploaded_file.name))
            
            else:
                df_entrada = pd.read_excel(uploaded_file, engine=_get_engine(uploaded_file.name))
    except Exception as e:
        st.error(f"❌ Erro ao ler o arquivo: {e}")
        st.stop()
    
    st.success(f"✅ Arquivo carregado: **{uploaded_file.name}** ({len(df_entrada)} registros, {len(df_entrada.columns)} colunas)")
    
    # Mostrar preview dos dados de entrada
    with st.expander("👁️ Preview dos dados de entrada", expanded=False):
        st.dataframe(df_entrada.head(10), use_container_width=True)
    
    st.markdown("---")
    
    # ============================================================
    # MAPEAMENTO DE COLUNAS
    # ============================================================
    st.subheader("🗺️ Mapeamento de Colunas")
    st.markdown(f"**Padrão:** {tipo_label} | **Origem:** {origem_label}")
    
    # Carregar colunas do modelo
    try:
        colunas_padrao = carregar_modelo(tipo)
    except Exception as e:
        st.error(f"❌ Erro ao carregar modelo: {e}")
        st.stop()
    
    colunas_origem = ['(deixar vazio)'] + list(df_entrada.columns)
    
    # Calcular mapeamento sugerido
    if origem in MAPEAMENTOS and tipo in MAPEAMENTOS[origem]:
        mapeamento_definido = MAPEAMENTOS[origem][tipo]
        mapeamento_sugerido = {}
        for col_padrao, opcoes in mapeamento_definido.items():
            col_encontrada = _encontrar_coluna(list(df_entrada.columns), opcoes)
            if col_encontrada:
                mapeamento_sugerido[col_padrao] = col_encontrada
        
        auto = _mapear_automatico(df_entrada, [c for c in colunas_padrao if c not in mapeamento_sugerido])
        mapeamento_sugerido.update(auto)
    else:
        mapeamento_sugerido = _mapear_automatico(df_entrada, colunas_padrao)
    
    # Interface de mapeamento editável
    st.markdown("Revise e ajuste o mapeamento abaixo. Cada coluna do padrão pode ser associada a uma coluna do arquivo de origem:")
    
    mapeamento_final = {}
    
    # Dividir em colunas para melhor visualização
    col1, col2 = st.columns(2)
    meio = len(colunas_padrao) // 2
    
    with col1:
        for col_padrao in colunas_padrao[:meio]:
            sugerido = mapeamento_sugerido.get(col_padrao)
            idx = colunas_origem.index(sugerido) if sugerido and sugerido in colunas_origem else 0
            
            escolha = st.selectbox(
                f"**{col_padrao}**",
                colunas_origem,
                index=idx,
                key=f"map_{col_padrao}"
            )
            if escolha != '(deixar vazio)':
                mapeamento_final[col_padrao] = escolha
    
    with col2:
        for col_padrao in colunas_padrao[meio:]:
            sugerido = mapeamento_sugerido.get(col_padrao)
            idx = colunas_origem.index(sugerido) if sugerido and sugerido in colunas_origem else 0
            
            escolha = st.selectbox(
                f"**{col_padrao}**",
                colunas_origem,
                index=idx,
                key=f"map_{col_padrao}"
            )
            if escolha != '(deixar vazio)':
                mapeamento_final[col_padrao] = escolha
    
    # Resumo do mapeamento
    mapeadas = len(mapeamento_final)
    total = len(colunas_padrao)
    st.markdown(f"**Resumo:** {mapeadas}/{total} colunas mapeadas")
    
    # Colunas da origem não utilizadas
    usadas = set(mapeamento_final.values())
    nao_usadas = [c for c in df_entrada.columns if c not in usadas]
    if nao_usadas:
        with st.expander(f"⚠️ {len(nao_usadas)} colunas da origem não mapeadas"):
            st.write(nao_usadas)
    
    st.markdown("---")
    
    # ============================================================
    # PREENCHIMENTO MANUAL DE CAMPOS
    # ============================================================
    st.subheader("✏️ Preenchimento Manual")
    st.markdown("Defina valores fixos para campos que o cliente não preenche. Esses valores serão aplicados a **todos** os registros.")
    
    valores_manuais = {}
    
    # Campos com opções pré-definidas (listas do DOit)
    from conversor import _carregar_padronizacoes
    padronizacoes = _carregar_padronizacoes()
    
    if tipo == 'financeiro':
        # TIPO DE RECEITA/DESPESA
        opcoes_tipo_rd = ['(não preencher)'] + [f"{chave} — {descs[0]}" for chave, descs in TIPOS_RECEITA_DESPESA.items()]
        escolha_tipo_rd = st.selectbox(
            "📑 **TIPO DE RECEITA/DESPESA**",
            opcoes_tipo_rd,
            index=0,
            help="Selecione o tipo para aplicar a todos os lançamentos"
        )
        if escolha_tipo_rd != '(não preencher)':
            valores_manuais['TIPO DE RECEITA/DESPESA'] = escolha_tipo_rd.split(' — ')[0]
        
        # DEPARTAMENTO
        if 'departamentos_financeiro' in padronizacoes:
            opcoes_dept = ['(não preencher)'] + padronizacoes['departamentos_financeiro']
            escolha_dept = st.selectbox("🏢 **DEPARTAMENTO**", opcoes_dept, index=0)
            if escolha_dept != '(não preencher)':
                valores_manuais['DEPARTAMENTO'] = escolha_dept
        
        # TIPO (Receita/Despesa)
        opcoes_tipo = ['(não preencher)', 'Receita', 'Despesa']
        escolha_tipo = st.selectbox("💰 **TIPO** (Receita/Despesa)", opcoes_tipo, index=0)
        if escolha_tipo != '(não preencher)':
            valores_manuais['TIPO'] = escolha_tipo
        
        # CONCILIADO
        opcoes_conc = ['(não preencher)', 'Sim', 'Não']
        escolha_conc = st.selectbox("✅ **CONCILIADO**", opcoes_conc, index=0)
        if escolha_conc != '(não preencher)':
            valores_manuais['CONCILIADO'] = escolha_conc
        
        # 1ª CATEGORIA (Nível 1 do plano de contas)
        if 'categorias_financeiro' in padronizacoes:
            nivel1 = sorted(set(l.split('>')[0].strip() for l in padronizacoes['categorias_financeiro'] if '>' in l))
            opcoes_cat1 = ['(não preencher)'] + nivel1
            escolha_cat1 = st.selectbox("📊 **1ª CATEGORIA**", opcoes_cat1, index=0)
            if escolha_cat1 != '(não preencher)':
                valores_manuais['1ª CATEGORIA'] = escolha_cat1
                
                # 2ª CATEGORIA (filtrada pelo nível 1)
                nivel2 = sorted(set(
                    l.split('>')[1].strip() 
                    for l in padronizacoes['categorias_financeiro'] 
                    if '>' in l and l.split('>')[0].strip() == escolha_cat1 and l.split('>')[1].strip()
                ))
                if nivel2:
                    opcoes_cat2 = ['(não preencher)'] + nivel2
                    escolha_cat2 = st.selectbox("📊 **2ª CATEGORIA**", opcoes_cat2, index=0)
                    if escolha_cat2 != '(não preencher)':
                        valores_manuais['2ª CATEGORIA'] = escolha_cat2
                        
                        # 3ª CATEGORIA (filtrada pelo nível 2)
                        nivel3 = sorted(set(
                            l.split('>')[2].strip()
                            for l in padronizacoes['categorias_financeiro']
                            if '>' in l and len(l.split('>')) >= 3
                            and l.split('>')[0].strip() == escolha_cat1
                            and l.split('>')[1].strip() == escolha_cat2
                            and l.split('>')[2].strip()
                        ))
                        if nivel3:
                            opcoes_cat3 = ['(não preencher)'] + nivel3
                            escolha_cat3 = st.selectbox("📊 **3ª CATEGORIA**", opcoes_cat3, index=0)
                            if escolha_cat3 != '(não preencher)':
                                valores_manuais['3ª CATEGORIA'] = escolha_cat3
    
    elif tipo == 'contatos':
        # CLASSIFICAÇÃO
        if 'classificacoes_cadastro' in padronizacoes:
            opcoes_class = ['(não preencher)'] + padronizacoes['classificacoes_cadastro']
            escolha_class = st.selectbox("🏷️ **CLASSIFICAÇÃO**", opcoes_class, index=0)
            if escolha_class != '(não preencher)':
                valores_manuais['CLASSIFICAÇÃO'] = escolha_class
        
        # TIPO DE ENDEREÇO
        if 'tipos_endereco' in padronizacoes:
            opcoes_end = ['(não preencher)'] + padronizacoes['tipos_endereco']
            escolha_end = st.selectbox("📍 **TIPO DE ENDEREÇO 1**", opcoes_end, index=0)
            if escolha_end != '(não preencher)':
                valores_manuais['TIPO DE ENDEREÇO 1'] = escolha_end
        
        # FORMA DE PAGAMENTO
        if 'formas_pagamento' in padronizacoes:
            opcoes_fp = ['(não preencher)'] + padronizacoes['formas_pagamento']
            escolha_fp = st.selectbox("💳 **FORMA DE PAGAMENTO**", opcoes_fp, index=0)
            if escolha_fp != '(não preencher)':
                valores_manuais['FORMA DE PAGAMENTO'] = escolha_fp
        
        # PAÍS
        opcoes_pais = ['(não preencher)', 'Brasil']
        escolha_pais = st.selectbox("🌍 **PAÍS 1**", opcoes_pais, index=0)
        if escolha_pais != '(não preencher)':
            valores_manuais['PAÍS 1'] = escolha_pais
    
    elif tipo == 'projetos':
        # CATEGORIA
        if 'categorias_projeto' in padronizacoes:
            opcoes_cat = ['(não preencher)'] + padronizacoes['categorias_projeto']
            escolha_cat = st.selectbox("📂 **CATEGORIA**", opcoes_cat, index=0)
            if escolha_cat != '(não preencher)':
                valores_manuais['CATEGORIA'] = escolha_cat
        
        # STATUS
        opcoes_status = ['(não preencher)', 'Ativo', 'Inativo', 'Concluído', 'Em andamento', 'Planejamento']
        escolha_status = st.selectbox("📋 **STATUS**", opcoes_status, index=0)
        if escolha_status != '(não preencher)':
            valores_manuais['STATUS'] = escolha_status
    
    elif tipo == 'horas':
        # STATUS
        opcoes_status_h = ['(não preencher)', 'Aprovado', 'Pendente', 'Rejeitado']
        escolha_status_h = st.selectbox("📋 **STATUS**", opcoes_status_h, index=0)
        if escolha_status_h != '(não preencher)':
            valores_manuais['STATUS'] = escolha_status_h
    
    # Mostrar resumo dos valores manuais
    if valores_manuais:
        st.info(f"📝 **{len(valores_manuais)} campo(s) serão preenchidos manualmente:** {', '.join(f'{k}={v}' for k,v in valores_manuais.items())}")
    
    st.markdown("---")
    
    # ============================================================
    # CONVERSÃO
    # ============================================================
    if st.button("🚀 Converter Arquivo", type="primary", use_container_width=True):
        with st.spinner("Convertendo..."):
            # Carregar referências se fornecidas
            df_cadastro_ref = None
            df_projetos_ref = None
            if cadastro_ref_file is not None:
                df_cadastro_ref = pd.read_excel(cadastro_ref_file)
            if projetos_ref_file is not None:
                df_projetos_ref = pd.read_excel(projetos_ref_file)
            
            # Construir DataFrame de saída
            df_saida = pd.DataFrame(columns=colunas_padrao)
            
            for col_padrao in colunas_padrao:
                if col_padrao in mapeamento_final and mapeamento_final[col_padrao] in df_entrada.columns:
                    df_saida[col_padrao] = df_entrada[mapeamento_final[col_padrao]].values
                else:
                    df_saida[col_padrao] = ''
            
            # ID sequencial
            if 'ID' in df_saida.columns:
                df_saida['ID'] = range(id_inicial, id_inicial + len(df_saida))
            
            # Aplicar valores manuais (campos preenchidos pelo usuário)
            for campo, valor in valores_manuais.items():
                if campo in df_saida.columns:
                    # Só preencher onde está vazio
                    mask = df_saida[campo].apply(lambda x: pd.isna(x) or str(x).strip() == '')
                    df_saida.loc[mask, campo] = valor
            
            # Campos extras → Anotações
            if tipo == 'contatos' and 'ANOTAÇÕES' in df_saida.columns:
                from conversor import _campos_extras_para_anotacoes
                anotacoes_extras = _campos_extras_para_anotacoes(df_entrada, mapeamento_final, colunas_padrao)
                existente = df_saida['ANOTAÇÕES'].apply(lambda x: str(x).strip() if pd.notna(x) and str(x).strip() else '')
                df_saida['ANOTAÇÕES'] = existente.combine(anotacoes_extras,
                    lambda a, b: f"{a} | {b}" if a and b else (a or b))
            
            # Vincular IDs
            if tipo == 'projetos' and df_cadastro_ref is not None and 'ID DO CADASTRO' in df_saida.columns:
                if 'NOME' in df_cadastro_ref.columns:
                    mapa_nome_id = dict(zip(
                        df_cadastro_ref['NOME'].str.lower().str.strip(),
                        df_cadastro_ref['ID']
                    ))
                    from conversor import _encontrar_coluna
                    col_cliente_origem = _encontrar_coluna(
                        list(df_entrada.columns),
                        ['Cliente', 'cliente', 'CLIENTE', 'Card Name', 'Nome do Cliente', 'Razão Social', 'Nome']
                    )
                    if col_cliente_origem:
                        df_saida['ID DO CADASTRO'] = df_entrada[col_cliente_origem].apply(
                            lambda x: mapa_nome_id.get(str(x).lower().strip(), '') if pd.notna(x) else ''
                        )
            
            if tipo == 'financeiro':
                if df_cadastro_ref is not None and 'ID DE / PARA' in df_saida.columns:
                    if 'NOME' in df_cadastro_ref.columns:
                        mapa_nome_id = dict(zip(
                            df_cadastro_ref['NOME'].str.lower().str.strip(),
                            df_cadastro_ref['ID']
                        ))
                        from conversor import _encontrar_coluna
                        col_cliente_origem = _encontrar_coluna(
                            list(df_entrada.columns),
                            ['Cliente', 'cliente', 'Fornecedor', 'Nome', 'Razão Social', 'DE / PARA', 'De/Para']
                        )
                        if col_cliente_origem:
                            df_saida['ID DE / PARA'] = df_entrada[col_cliente_origem].apply(
                                lambda x: mapa_nome_id.get(str(x).lower().strip(), '') if pd.notna(x) else ''
                            )
                
                if df_projetos_ref is not None and 'ID PROJETO' in df_saida.columns:
                    if 'NOME' in df_projetos_ref.columns:
                        mapa_proj_id = dict(zip(
                            df_projetos_ref['NOME'].str.lower().str.strip(),
                            df_projetos_ref['ID']
                        ))
                        from conversor import _encontrar_coluna
                        col_projeto_origem = _encontrar_coluna(
                            list(df_entrada.columns),
                            ['Projeto', 'projeto', 'PROJETO', 'Obra', 'Nome do Projeto', 'Project']
                        )
                        if col_projeto_origem:
                            df_saida['ID PROJETO'] = df_entrada[col_projeto_origem].apply(
                                lambda x: mapa_proj_id.get(str(x).lower().strip(), '') if pd.notna(x) else ''
                            )
            
            # Aplicar formatações
            df_saida = _aplicar_formatacoes(df_saida, tipo)
            
            # Aplicar estilo de caixa
            from conversor import _aplicar_caixa, _validar_padronizacoes, _carregar_padronizacoes
            df_saida = _aplicar_caixa(df_saida, estilo_caixa)
            
            # Validar padronizações
            padronizacoes = _carregar_padronizacoes()
            alertas = _validar_padronizacoes(df_saida, tipo, padronizacoes)
        
        st.success(f"✅ Conversão concluída! {len(df_saida)} registros convertidos.")
        
        # Mostrar alertas de padronização
        if alertas:
            st.warning("⚠️ **Valores novos encontrados (não existem no DOit):**")
            for alerta in alertas:
                st.markdown(f"- {alerta}")
            st.info("Você precisará criar esses itens no DOit antes de importar.")
        
        # Preview do resultado
        st.subheader("📊 Preview do Resultado")
        st.dataframe(df_saida.head(20), use_container_width=True)
        
        # Estatísticas
        col_stat1, col_stat2, col_stat3 = st.columns(3)
        with col_stat1:
            st.metric("Registros", len(df_saida))
        with col_stat2:
            preenchidas = df_saida.replace('', pd.NA).notna().sum().sum()
            total_celulas = df_saida.shape[0] * df_saida.shape[1]
            pct = (preenchidas / total_celulas * 100) if total_celulas > 0 else 0
            st.metric("Preenchimento", f"{pct:.1f}%")
        with col_stat3:
            st.metric("Colunas", len(df_saida.columns))
        
        # Download
        st.markdown("---")
        
        # Gerar Excel completo com abas auxiliares
        from abas_auxiliares import gerar_excel_completo
        
        # Carregar todas as abas do arquivo original (para detectar dados bancários, plano de contas)
        todas_abas_original = None
        if uploaded_file.name.endswith(('.xlsx', '.xls')):
            try:
                uploaded_file.seek(0)
                todas_abas_original = pd.read_excel(uploaded_file, engine=_get_engine(uploaded_file.name), sheet_name=None)
            except Exception:
                pass
        
        # Colunas não mapeadas
        usadas_final = set(mapeamento_final.values())
        nao_mapeadas_final = [c for c in df_entrada.columns if c not in usadas_final]
        
        excel_bytes = gerar_excel_completo(
            df_saida=df_saida,
            tipo=tipo,
            alertas=alertas,
            df_entrada=df_entrada,
            todas_abas=todas_abas_original,
            campos_nao_mapeados=nao_mapeadas_final,
        )
        
        nome_saida = f"{os.path.splitext(uploaded_file.name)[0]}_padrao_{tipo}.xlsx"
        
        st.download_button(
            label="📥 Baixar arquivo convertido (.xlsx)",
            data=excel_bytes,
            file_name=nome_saida,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

else:
    # Estado inicial - mostrar informações dos modelos
    st.info("👆 Faça upload de um arquivo para começar a conversão.")
    
    st.markdown("---")
    st.subheader("📋 Layouts Padrão Disponíveis")
    
    tab1, tab2, tab3, tab4 = st.tabs(["Contatos", "Projetos", "Financeiro", "Horas"])
    
    modelos_info = {
        'contatos': tab1,
        'projetos': tab2,
        'financeiro': tab3,
        'horas': tab4,
    }
    
    for tipo_info, tab in modelos_info.items():
        with tab:
            try:
                colunas = carregar_modelo(tipo_info)
                st.markdown(f"**{len(colunas)} colunas:**")
                cols_display = st.columns(3)
                for i, col in enumerate(colunas):
                    with cols_display[i % 3]:
                        st.markdown(f"• {col}")
            except Exception as e:
                st.warning(f"Modelo não disponível: {e}")
