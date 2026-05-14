"""
Módulo Conversor de Dados para Implantação
Converte arquivos de diferentes sistemas (ClickUp, Sienge, Trello, Excel manual)
para o padrão DOit (planilhas modelo).

Uso via linha de comando:
    python conversor.py --arquivo entrada.xlsx --tipo contatos --origem clickup --saida resultado.xlsx

Uso como módulo:
    from conversor import converter_arquivo
    df_resultado = converter_arquivo("entrada.xlsx", tipo="contatos", origem="clickup")
"""

import os
import sys
import argparse
import pandas as pd
import warnings

warnings.filterwarnings('ignore')

# Diretório das planilhas padrão
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MODELOS_DIR = os.path.join(BASE_DIR, '..', 'models')

# ============================================================
# TIPOS DE RECEITA/DESPESA (chave DOit ← descrição amigável)
# ============================================================
TIPOS_RECEITA_DESPESA = {
    'ADMINISTRATIVE_TAX': ['Taxa Administrativa', 'Taxa Admin', 'Administrativa'],
    'EMPLOYEE_REFUND': ['Reembolso de Funcionário', 'Reembolso Funcionário', 'Reembolso Func'],
    'EXPENSE_REFUND': ['Reembolso de Despesa', 'Reembolso Despesa', 'Reembolso'],
    'OVERTIME': ['Hora Extra', 'Horas Extras', 'HE'],
    'PROJECT_EXPENSE': ['Despesa de Projeto', 'Despesa Projeto', 'Desp. Projeto'],
    'PROJECT_INVOICE': ['Pedido (Projeto)', 'Pedido Projeto', 'Pedido', 'Invoice Projeto'],
    'PROJECT_PHASE': ['Etapa de Projeto', 'Etapa Projeto', 'Etapa', 'Parcela'],
    'RESERVE': ['Reserva Técnica', 'Reserva', 'RT'],
    'SALE_INVOICE': ['Faturamento', 'Fatura', 'NF', 'Nota Fiscal'],
    'SERVICE_INVOICE': ['Serviço', 'Servico', 'NFS', 'Nota de Serviço'],
    'THIRD_PARTY': ['Terceiros', 'Terceiro', '3os'],
    'VISITS': ['Visitas', 'Visita', 'Visita Técnica'],
}

# Mapa reverso: descrição → chave
_MAPA_TIPO_REVERSO = {}
for chave, descricoes in TIPOS_RECEITA_DESPESA.items():
    _MAPA_TIPO_REVERSO[chave.lower()] = chave  # a própria chave
    for desc in descricoes:
        _MAPA_TIPO_REVERSO[desc.lower()] = chave


def _traduzir_tipo_receita_despesa(valor) -> str:
    """Traduz um valor de tipo de receita/despesa para a chave DOit."""
    if pd.isna(valor) or str(valor).strip() == '':
        return ''
    val = str(valor).strip()
    # Tentar match direto (case insensitive)
    if val.lower() in _MAPA_TIPO_REVERSO:
        return _MAPA_TIPO_REVERSO[val.lower()]
    # Tentar match parcial
    for desc_lower, chave in _MAPA_TIPO_REVERSO.items():
        if desc_lower in val.lower() or val.lower() in desc_lower:
            return chave
    # Se não encontrou, retorna o valor original (vai gerar alerta)
    return val

# ============================================================
# LAYOUTS PADRÃO (colunas esperadas por tipo)
# ============================================================

def carregar_modelo(tipo: str) -> list:
    """Carrega as colunas do modelo padrão a partir da planilha."""
    mapa_arquivos = {
        'contatos': 'doit-modelo-contatos.xlsx',
        'projetos': 'doit-modelo-projetos.xlsx',
        'financeiro': 'doit-modelo-financeiro.xlsx',
        'horas': 'doit-modelo-horas-trabalhadas.xlsx',
        'usuarios': 'doit-modelo-usuarios.xlsx',
        'produtos': 'doit-modelo-produtos.xlsx',
        'vendas': 'doit-modelo-venda.xlsx',
    }
    
    arquivo = mapa_arquivos.get(tipo)
    if not arquivo:
        raise ValueError(f"Tipo '{tipo}' não reconhecido. Use: {list(mapa_arquivos.keys())}")
    
    caminho = os.path.join(MODELOS_DIR, arquivo)
    if not os.path.exists(caminho):
        raise FileNotFoundError(f"Modelo não encontrado: {caminho}")
    
    df_modelo = pd.read_excel(caminho, nrows=0)
    return list(df_modelo.columns)


# ============================================================
# MAPEAMENTOS POR SISTEMA DE ORIGEM
# ============================================================

# Cada mapeamento é um dicionário: {coluna_padrao: coluna_origem}
# Se a coluna_origem for None, será preenchida em branco
# Se for uma lista, tenta cada opção em ordem

MAPEAMENTOS = {
    'clickup': {
        'contatos': {
            'NOME': ['Name', 'Nome', 'name', 'ASSIGNEES'],
            'EMAIL': ['Email', 'email', 'E-mail', 'EMAIL'],
            'TELEFONE COMERCIAL': ['Phone', 'Telefone', 'phone'],
            'CELULAR': ['Mobile', 'Celular', 'celular'],
            'DOCUMENTO FEDERAL': ['CPF', 'CNPJ', 'CPF/CNPJ', 'Documento'],
            'ANOTAÇÕES': ['Description', 'Descrição', 'Notes'],
            'WEBSITE': ['Website', 'URL', 'Site'],
        },
        'projetos': {
            'NOME': ['Task Name', 'Name', 'Nome', 'TASK NAME'],
            'DESCRIÇÃO': ['Description', 'Descrição', 'DESCRIPTION'],
            'CATEGORIA': ['List Name', 'Lista', 'Category', 'Tags'],
            'STATUS': ['Status', 'STATUS', 'status'],
            'INÍCIO': ['Start Date', 'Data Início', 'START DATE', 'Created'],
            'TÉRMINO': ['Due Date', 'Data Fim', 'DUE DATE', 'End Date'],
        },
        'horas': {
            'DATA INICIAL': ['Date', 'Data', 'Start Date', 'DATE'],
            'HORÁRIO INICIAL': ['Start Time', 'Hora Início'],
            'DATA FINAL': ['End Date', 'Data Final'],
            'HORÁRIO FINAL': ['End Time', 'Hora Fim'],
            'ID DO COLABORADOR': ['Assignee', 'Responsável', 'ASSIGNEES', 'User'],
            'NOME DA ATIVIDADE': ['Task Name', 'Nome da Tarefa', 'TASK NAME'],
            'DESCRIÇÃO DA ATIVIDADE': ['Description', 'Descrição', 'Note'],
            'HORAS TRABALHADAS': ['Duration', 'Duração', 'Hours', 'Time Tracked', 'TIME TRACKED'],
        },
        'financeiro': {
            'DATA': ['Baixa', 'Date', 'Data', 'DATE'],
            'EMISSÃO': ['Emissão', 'Created Date'],
            'VENCIMENTO': ['Vencto', 'Due Date', 'Vencimento', 'DUE DATE'],
            'DESCRIÇÃO': ['Descr. Baixa', 'Observações', 'Task Name', 'Description', 'Descrição', 'TASK NAME'],
            'VALOR': ['Valor', 'V. Bruto', 'Value', 'Amount', 'CUSTOM FIELD: Valor'],
            'DI': ['Nº Doc.', 'Documento'],
            'TIPO': ['R/D', 'Type', 'Tipo', 'Tipo Baixa', 'Category'],
            'ID DE / PARA': ['Favorecido / Sacado', 'Favorecido', 'Sacado'],
            '1ª CATEGORIA': ['Classificação Financeira', 'Classificação', 'Category'],
            'CONCILIADO': ['Tipo Baixa'],
            'TIPO DE RECEITA/DESPESA': ['Tipo Receita', 'Tipo Despesa', 'Receita/Despesa', 'AccountingEntryType'],
        },
        'usuarios': {},
        'produtos': {},
        'vendas': {},
    },
    'sienge': {
        'contatos': {
            'ID': ['Código', 'Codigo', 'codigo', 'ID'],
            'NOME': ['Nome', 'Razão Social', 'NOME', 'razaoSocial'],
            'APELIDO': ['Nome fantasia', 'Nome Fantasia', 'Apelido', 'nomeFantasia'],
            'EMAIL': ['E-mail', 'Email', 'email'],
            'TELEFONE COMERCIAL': ['Telefone principal', 'Telefone', 'Fone Comercial', 'telefone'],
            'CELULAR': ['Celular', 'celular'],
            'RUA 1': ['Endereço', 'Logradouro', 'Rua', 'endereco'],
            'NÚMERO 1': ['Número do endereço', 'Número', 'Nro', 'numero'],
            'COMPLEMENTO 1': ['Complemento', 'complemento'],
            'BAIRRO 1': ['Bairro', 'bairro'],
            'CIDADE 1': ['Município', 'Cidade', 'cidade'],
            'ESTADO 1': ['UF', 'Estado', 'uf'],
            'CEP 1': ['CEP', 'cep'],
            'DOCUMENTO FEDERAL': ['CNPJ/CPF', 'CNPJ', 'CPF', 'CPF/CNPJ', 'cnpj', 'cpf'],
            'DOCUMENTO ESTADUAL': ['Inscrição estadual', 'IE', 'Inscrição Estadual', 'inscricaoEstadual'],
            'CLASSIFICAÇÃO': ['Descrição tipo cliente', 'Tipo pessoa', 'Classificação'],
            'TIPO': ['Tipo pessoa', 'Tipo', 'tipo'],
        },
        'projetos': {
            'NOME': ['Descrição', 'Nome da Obra', 'Obra', 'descricao'],
            'DESCRIÇÃO': ['Observação', 'Obs', 'observacao'],
            'CATEGORIA': ['Tipo', 'Tipo de Obra', 'tipo'],
            'STATUS': ['Situação', 'Status', 'situacao'],
            'INÍCIO': ['Data Início', 'Início', 'dataInicio'],
            'TÉRMINO': ['Data Término', 'Previsão Término', 'dataTermino'],
        },
        'financeiro': {
            'DATA': ['Data', 'Data Lançamento', 'dataLancamento', 'Baixa'],
            'EMISSÃO': ['Data Emissão', 'Emissão', 'dataEmissao'],
            'VENCIMENTO': ['Data Vencimento', 'Vencimento', 'dataVencimento'],
            'DESCRIÇÃO': ['Histórico', 'Descrição', 'descricao'],
            'VALOR': ['Crédito', 'Valor', 'Valor Total', 'valor'],
            'DI': ['Documento', 'Nº Doc.', 'Título/Parcela'],
            'TIPO': ['Tipo', 'Natureza', 'tipo', 'OR'],
            'DEPARTAMENTO': ['Centro de Custo', 'Departamento', 'centroCusto'],
            '1ª CATEGORIA': ['Categoria', 'Plano de Contas', 'categoria', 'Classificação Financeira'],
            'TIPO DE RECEITA/DESPESA': ['Tipo Receita', 'Tipo Despesa', 'Receita/Despesa', 'tipoLancamento'],
            'ID DE / PARA': ['Favorecido / Sacado', 'Favorecido', 'Sacado', 'Cliente', 'Fornecedor'],
        },
        'horas': {
            'DATA INICIAL': ['Data', 'Data Início', 'data'],
            'ID DO COLABORADOR': ['Funcionário', 'Colaborador', 'Nome', 'funcionario'],
            'NOME DA ATIVIDADE': ['Atividade', 'Tarefa', 'Serviço', 'atividade'],
            'DESCRIÇÃO DA ATIVIDADE': ['Observação', 'Descrição', 'observacao'],
            'HORAS TRABALHADAS': ['Horas', 'Total Horas', 'Quantidade', 'horas'],
        },
        'usuarios': {},
        'produtos': {},
        'vendas': {},
    },
    'trello': {
        'contatos': {
            'NOME': ['Card Name', 'Nome', 'name'],
            'EMAIL': ['Email', 'E-mail'],
            'ANOTAÇÕES': ['Card Description', 'Description', 'Descrição'],
        },
        'projetos': {
            'NOME': ['Card Name', 'Nome', 'name'],
            'DESCRIÇÃO': ['Card Description', 'Description', 'Descrição'],
            'CATEGORIA': ['List Name', 'Lista', 'Board Name'],
            'STATUS': ['List Name', 'Status'],
            'INÍCIO': ['Start Date', 'Data Início', 'Created Date'],
            'TÉRMINO': ['Due Date', 'Data Fim'],
        },
        'horas': {
            'DATA INICIAL': ['Date', 'Data', 'Start Date'],
            'ID DO COLABORADOR': ['Members', 'Responsável', 'Assignee'],
            'NOME DA ATIVIDADE': ['Card Name', 'Nome'],
            'DESCRIÇÃO DA ATIVIDADE': ['Card Description', 'Description'],
            'HORAS TRABALHADAS': ['Hours', 'Horas', 'Duration'],
        },
        'financeiro': {
            'DATA': ['Date', 'Data'],
            'DESCRIÇÃO': ['Card Name', 'Nome', 'Description'],
            'VALOR': ['Value', 'Valor', 'Amount'],
            'TIPO': ['List Name', 'Labels', 'Tipo'],
            'VENCIMENTO': ['Due Date', 'Vencimento'],
        },
        'usuarios': {},
        'produtos': {},
        'vendas': {},
    },
    'outlook': {
        'contatos': {
            'NOME': ['Primeiro nome', 'Nome para Exibição do Email'],
            'APELIDO': ['Empresa'],
            'EMAIL': ['E-mail Address', 'Endereço de email 2', 'Endereço de email 3'],
            'TELEFONE COMERCIAL': ['Telefone Comercial', 'Telefone Comercial 2', 'Telefone principal da empresa', 'Telefone principal'],
            'TELEFONE RESIDENCIAL': ['Telefone residencial', 'Telefone residencial 2'],
            'CELULAR': ['Telefone celular'],
            'FAX': ['Fax Comercial', 'Fax residencial'],
            'RUA 1': ['Business Street', 'Endereço residencial'],
            'COMPLEMENTO 1': ['Rua do endereço comercial 2', 'Endereço residencial 2'],
            'BAIRRO 1': ['Rua do endereço comercial 3', 'Endereço residencial 3'],
            'CIDADE 1': ['Business City', 'Cidade do endereço residencial'],
            'ESTADO 1': ['Business State', 'Estado'],
            'CEP 1': ['Business Postal Code', 'CEP do endereço residencial'],
            'PAÍS 1': ['País/Região da Empresa', 'País/Região de Residência'],
            'ANOTAÇÕES': ['Anotações'],
            'WEBSITE': ['Página da Web'],
            'CLASSIFICAÇÃO': ['Categorias'],
        },
        'projetos': {},
        'financeiro': {},
        'horas': {},
        'usuarios': {},
        'produtos': {},
        'vendas': {},
    },
    'excel_manual': {
        # Para Excel manual, tentamos mapear por similaridade de nomes
        'contatos': {},
        'projetos': {},
        'financeiro': {},
        'horas': {},
        'usuarios': {},
        'produtos': {},
        'vendas': {},
    },
    'financeiro_horizontal': {
        'contatos': {},
        'projetos': {},
        'financeiro': {
            'DATA': ['DATA'],
            'VENCIMENTO': ['VENCIMENTO'],
            'DESCRIÇÃO': ['DESCRIÇÃO'],
            'VALOR': ['VALOR'],
            'TIPO': ['TIPO'],
            'CONCILIADO': ['CONCILIADO'],
            '1ª CATEGORIA': ['1ª CATEGORIA'],
            '2ª CATEGORIA': ['2ª CATEGORIA'],
            '3ª CATEGORIA': ['3ª CATEGORIA'],
        },
        'horas': {},
        'usuarios': {},
        'produtos': {},
        'vendas': {},
    },
    'excel_desestruturado': {
        'contatos': {
            'NOME': ['Nome', 'NOME', 'name', 'Cliente', 'CLIENTE', 'Razão Social'],
            'EMAIL': ['Email', 'E-mail', 'email', 'EMAIL'],
            'TELEFONE COMERCIAL': ['Telefone', 'Tel', 'Fone', 'TELEFONE'],
            'CELULAR': ['Celular', 'Cel', 'WhatsApp', 'Whats'],
        },
        'projetos': {
            'NOME': ['Nome', 'NOME', 'Projeto', 'PROJETO', 'name', 'Card Name', 'Cliente'],
            'DESCRIÇÃO': ['Descrição', 'DESCRIÇÃO', 'DESCRIÇÃO  DE ANDAMENTO', 'Obs', 'Observação', 'Description'],
            'CATEGORIA': ['Categoria', 'CATEGORIA', 'Tipo', 'NÍVEL', 'Status', 'NÍVEL'],
            'STATUS': ['Status', 'STATUS', 'Situação', 'NÍVEL'],
            'INÍCIO': ['Data', 'DATA DE INICIO', 'DATA DE INICIO PRJ', 'Início', 'Start Date', 'Data Início', 'INICIO PRJ'],
            'TÉRMINO': ['Término', 'Data Fim', 'Due Date', 'Entrega', 'DATA DE ENTREGA OBRA'],
            'ID LÍDER': ['Responsável', 'RESPONSÁVEL', 'Projetista', 'PROJETISTA', 'Gestão'],
        },
        'financeiro': {
            'DATA': ['Data', 'DATA', 'Date'],
            'DESCRIÇÃO': ['Descrição', 'DESCRIÇÃO', 'Description', 'Histórico'],
            'VALOR': ['Valor', 'VALOR', 'Value', 'Amount'],
            'TIPO': ['Tipo', 'TIPO', 'Type'],
            'VENCIMENTO': ['Vencimento', 'Data Vencimento', 'Due Date'],
        },
        'horas': {
            'DATA INICIAL': ['Data', 'DATA', 'Date'],
            'ID DO COLABORADOR': ['Responsável', 'Colaborador', 'Nome', 'Profissional'],
            'NOME DA ATIVIDADE': ['Atividade', 'Tarefa', 'Projeto', 'Nome'],
            'HORAS TRABALHADAS': ['Horas', 'Duração', 'Hours', 'Tempo'],
        },
        'usuarios': {},
        'produtos': {},
        'vendas': {},
    },
    'conta_azul': {
        'contatos': {
            'NOME': ['Nome', 'Nome Contato'],
            'APELIDO': ['Razão Social'],
            'EMAIL': ['Email principal', 'E-mail Contato'],
            'TELEFONE COMERCIAL': ['Telefone principal'],
            'CELULAR': ['Telefone principal'],
            'RUA 1': ['Endereço'],
            'NÚMERO 1': ['Número'],
            'COMPLEMENTO 1': ['Complemento'],
            'BAIRRO 1': ['Bairro'],
            'CIDADE 1': ['Cidade'],
            'ESTADO 1': ['Estado'],
            'CEP 1': ['CEP'],
            'DOCUMENTO FEDERAL': ['CNPJ', 'CPF'],
            'DOCUMENTO ESTADUAL': ['Inscrição Estadual'],
            'CLASSIFICAÇÃO': ['Status'],
            'ANOTAÇÕES': ['Observações'],
            'WEBSITE': ['Website'],
        },
        'projetos': {},
        'financeiro': {
            'DATA': ['Data movimento'],
            'EMISSÃO': ['Data de competência'],
            'VENCIMENTO': ['Data original de vencimento', 'Data prevista'],
            'DESCRIÇÃO': ['Descrição'],
            'VALOR': ['Valor (R$)', 'Valor original (R$)'],
            'DI': ['Identificador do fornecedor/cliente'],
            'TIPO': ['Tipo da operação'],
            'CONCILIADO': ['Situação'],
            'ID DE / PARA': ['Nome do fornecedor/cliente'],
            '1ª CATEGORIA': ['Categoria 1'],
            'DEPARTAMENTO': ['Centro de Custo 1'],
            'FORMA DE PAGAMENTO': ['Forma de pgto/recbto'],
            'TIPO DE RECEITA/DESPESA': ['Tipo da operação'],
        },
        'horas': {},
        'usuarios': {},
        'produtos': {},
        'vendas': {},
    },
    'navis': {
        'contatos': {
            'NOME': ['Nome'],
            'APELIDO': ['Nome Fantasia'],
            'EMAIL': ['Email'],
            'TELEFONE COMERCIAL': ['Telefone'],
            'RUA 1': ['Endereço'],
            'BAIRRO 1': ['Bairro'],
            'CIDADE 1': ['Cidade'],
            'ESTADO 1': ['Estado'],
            'CEP 1': ['CEP'],
            'DOCUMENTO FEDERAL': ['CNPJ/CPF'],
            'DOCUMENTO ESTADUAL': ['IE/RG'],
            'ANOTAÇÕES': ['Atividade', 'Contatos', 'Contato'],
            'COMPLEMENTO 1': ['Complemento'],
            'WEBSITE': ['Site'],
            'CLASSIFICAÇÃO': ['Atividade'],
        },
        'projetos': {
            'NOME': ['Nome do Projeto'],
            'DESCRIÇÃO': ['Observações'],
            'CATEGORIA': ['Tipo de Projeto'],
            'STATUS': ['Status'],
            'INÍCIO': ['Dt.Início'],
            'TÉRMINO': ['Dt.Término'],
            'ID DO CADASTRO': ['Nome do Cliente', 'Nome Fantasia'],
        },
        'financeiro': {
            # Financeiro CC
            'DATA': ['Baixa', 'Movimento', 'Data movimento'],
            'EMISSÃO': ['Emissão'],
            'VENCIMENTO': ['Vencimento', 'Bom Para'],
            'DESCRIÇÃO': ['Observações', 'Descr. / Nro. Cheque'],
            'VALOR': ['Valor', 'V. Bruto', 'Valor Liq.'],
            'DI': ['Nro. Doc.', 'Nº Doc.'],
            'TIPO': ['Tipo'],
            'CONCILIADO': ['Conciliado_Texto'],
            'ID DE / PARA': ['Pessoa', 'Favorecido / Sacado'],
            '1ª CATEGORIA': ['Categoria', 'Classificação Financeira'],
            'FORMA DE PAGAMENTO': ['Descr. Baixa', 'Baixa/Lancto.'],
        },
        'horas': {
            'DATA INICIAL': ['Data'],
            'ID DO COLABORADOR': ['Colaborador'],
            'NOME DA ATIVIDADE': ['Projeto'],
            'HORAS TRABALHADAS': ['Horas'],
        },
        'usuarios': {},
        'produtos': {},
        'vendas': {},
    },
    'doit_coleta': {
        'contatos': {
            'NOME': ['NOME'],
            'APELIDO': ['APELIDO'],
            'EMAIL': ['EMAIL'],
            'TELEFONE COMERCIAL': ['TELEFONE COMERCIAL'],
            'TELEFONE RESIDENCIAL': ['TELEFONE RESIDENCIAL'],
            'CELULAR': ['CELULAR'],
            'FAX': ['FAX'],
            'RUA 1': ['RUA 1'],
            'NÚMERO 1': ['NÚMERO 1'],
            'COMPLEMENTO 1': ['COMPLEMENTO 1'],
            'BAIRRO 1': ['BAIRRO 1'],
            'CIDADE 1': ['CIDADE 1'],
            'ESTADO 1': ['ESTADO 1'],
            'CEP 1': ['CEP 1'],
            'PAÍS 1': ['PAÍS 1'],
            'RUA 2': ['RUA 2'],
            'NÚMERO 2': ['NÚMERO 2'],
            'COMPLEMENTO 2': ['COMPLEMENTO 2'],
            'BAIRRO 2': ['BAIRRO 2'],
            'CIDADE 2': ['CIDADE 2'],
            'ESTADO 2': ['ESTADO 2'],
            'CEP 2': ['CEP 2'],
            'PAÍS 2': ['PAÍS 2'],
            'ANOTAÇÕES': ['ANOTAÇÕES'],
            'WEBSITE': ['WEBSITE'],
            'ORIGEM': ['ORIGEM'],
            'DOCUMENTO FEDERAL': ['DOCUMENTO FEDERAL'],
            'DOCUMENTO ESTADUAL': ['DOCUMENTO ESTADUAL'],
            'DATA 1': ['DATA 1'],
            'DATA 2': ['DATA 2'],
            'DATA 3': ['DATA 3'],
            'DATA 4': ['DATA 4'],
            'DATA 5': ['DATA 5'],
            'CLASSIFICAÇÃO': ['CLASSIFICAÇÃO'],
        },
        'projetos': {
            'NOME': ['NOME'],
            'DESCRIÇÃO': ['DESCRIÇÃO'],
            'ID DO CADASTRO': ['CLIENTE'],
            'CATEGORIA': ['CATEGORIA'],
            'STATUS': ['STATUS'],
            'INÍCIO': ['DATA CONTRATO'],
            'EXECUÇÃO': ['EXECUÇÃO'],
            'TÉRMINO': ['TÉRMINO'],
            'REUNIÕES CONTRATADAS (EM HORAS)': ['REUNIÕES CONTRATADAS (EM HORAS)'],
            'ATIVIDADES CONTRATADAS (EM HORAS)': ['ATIVIDADES CONTRATADAS (EM HORAS)'],
            'VISITAS CONTRATADAS (NÚMERO DE VISITAS)': ['VISITAS CONTRATADAS (NÚMERO DE VISITAS)'],
            'TAXA DE ADMINISTRAÇÃO': ['TAXA DE ADMINISTRAÇÃO'],
            'CUSTO VISITA': ['CUSTO VISITA'],
            'COLUNA CUSTOMIZÁVEL DATA 1': ['COLUNA DE DATA CUSTOMIZÁVEL'],
            'COLUNA CUSTOMIZÁVEL NUMERO 1': ['COLUNA DE NÚMERO CUSTOMIZÁVEL'],
            'COLUNA CUSTOMIZÁVEL BOOLEANA 1': ['COLUNA DE CHECK CUSTOMIZÁVEL'],
            'COLUNA CUSTOMIZÁVEL TEXTO 1': ['COLUNA DE ITENS CUSTOMIZÁVEL'],
        },
        'financeiro': {
            'DATA': ['DATA REALIZADO'],
            'EMISSÃO': ['DATA DE EMISSÃO'],
            'VENCIMENTO': ['DATA DE VENCIMENTO'],
            'DESCRIÇÃO': ['DESCRIÇÃO'],
            'DI': ['DOC. INT'],
            'DE': ['DOC. EXT'],
            'VALOR': ['VALOR'],
            'CONCILIADO': ['CONCILIADO'],
            'ID PROJETO': ['PROJETO'],
            'ID DE / PARA': ['DE / PARA'],
            'TIPO': ['TIPO'],
            'DEPARTAMENTO': ['DEPARTAMENTO'],
            '1ª CATEGORIA': ['1ª CATEGORIA'],
            '2ª CATEGORIA': ['2ª CATEGORIA'],
            '3ª CATEGORIA': ['3ª CATEGORIA'],
        },
        'horas': {},
        'usuarios': {
            'NOME': ['NOME', 'Nome'],
            'EMAIL': ['EMAIL', 'Email', 'E-mail'],
        },
        'produtos': {},
        'vendas': {},
    },
}


# ============================================================
# FUNÇÕES DE CONVERSÃO
# ============================================================

def _encontrar_coluna(df_colunas: list, opcoes: list) -> str:
    """Encontra a primeira coluna que existe no DataFrame."""
    if not opcoes:
        return None
    df_colunas_lower = {c.lower().strip(): c for c in df_colunas}
    for opcao in opcoes:
        if opcao in df_colunas:
            return opcao
        if opcao.lower().strip() in df_colunas_lower:
            return df_colunas_lower[opcao.lower().strip()]
    return None


def _mapear_automatico(df_origem: pd.DataFrame, colunas_padrao: list) -> dict:
    """
    Tenta mapear automaticamente colunas por similaridade de nome.
    Usado quando a origem é 'excel_manual' ou quando não há mapeamento definido.
    Cada coluna de origem só pode ser usada uma vez (evita duplicação).
    """
    mapeamento = {}
    colunas_origem = list(df_origem.columns)
    colunas_origem_lower = {c.lower().strip(): c for c in colunas_origem}
    usadas = set()
    
    # Primeira passada: match exato (case insensitive)
    for col_padrao in colunas_padrao:
        col_lower = col_padrao.lower().strip()
        if col_lower in colunas_origem_lower:
            col_real = colunas_origem_lower[col_lower]
            if col_real not in usadas:
                mapeamento[col_padrao] = col_real
                usadas.add(col_real)
    
    # Segunda passada: match parcial (apenas para colunas ainda não mapeadas)
    for col_padrao in colunas_padrao:
        if col_padrao in mapeamento:
            continue
        col_lower = col_padrao.lower().strip()
        for orig_lower, orig_real in colunas_origem_lower.items():
            if orig_real in usadas:
                continue
            if col_lower == orig_lower:
                mapeamento[col_padrao] = orig_real
                usadas.add(orig_real)
                break
    
    return mapeamento


def _formatar_telefone(valor):
    """Formata telefone para padrão Brasil: +55 (XX) XXXXX-XXXX ou +55 (XX) XXXX-XXXX."""
    if pd.isna(valor) or str(valor).strip() == '':
        return ''
    tel = ''.join(c for c in str(valor) if c.isdigit())
    # Remover código do país se já veio com 55 na frente
    if len(tel) == 13 and tel.startswith('55'):
        tel = tel[2:]
    elif len(tel) == 12 and tel.startswith('55'):
        tel = tel[2:]
    if len(tel) == 11:
        return f"+55 ({tel[:2]}) {tel[2:7]}-{tel[7:]}"
    elif len(tel) == 10:
        return f"+55 ({tel[:2]}) {tel[2:6]}-{tel[6:]}"
    elif len(tel) == 9:
        # Celular sem DDD - mantém sem DDD
        return f"{tel[:5]}-{tel[5:]}"
    elif len(tel) == 8:
        # Fixo sem DDD
        return f"{tel[:4]}-{tel[4:]}"
    return str(valor)


def _formatar_cpf_cnpj(valor):
    """Formata CPF ou CNPJ."""
    if pd.isna(valor) or str(valor).strip() == '':
        return ''
    doc = ''.join(c for c in str(valor) if c.isdigit())
    if len(doc) == 11:
        return f"{doc[:3]}.{doc[3:6]}.{doc[6:9]}-{doc[9:]}"
    elif len(doc) == 14:
        return f"{doc[:2]}.{doc[2:5]}.{doc[5:8]}/{doc[8:12]}-{doc[12:]}"
    return str(valor)


def _formatar_cep(valor):
    """Formata CEP para XXXXX-XXX."""
    if pd.isna(valor) or str(valor).strip() == '':
        return ''
    cep = ''.join(c for c in str(valor) if c.isdigit())
    if len(cep) == 8:
        return f"{cep[:5]}-{cep[5:]}"
    return str(valor)


def _converter_data(valor):
    """Tenta converter valor para data no formato DD/MM/YYYY."""
    if pd.isna(valor) or str(valor).strip() == '':
        return ''
    try:
        dt = pd.to_datetime(valor, dayfirst=True, errors='coerce')
        if pd.notna(dt):
            return dt.strftime('%d/%m/%Y')
    except Exception:
        pass
    return str(valor)


def _primeira_maiuscula(valor):
    """Converte texto para Primeira Maiúscula (Title Case), tratando preposições."""
    if pd.isna(valor) or str(valor).strip() == '':
        return ''
    texto = str(valor).strip().title()
    # Manter preposições em minúscula (exceto se for a primeira palavra)
    preposicoes = ['Da', 'De', 'Do', 'Das', 'Dos', 'E', 'Em', 'Para', 'Com', 'Por']
    palavras = texto.split()
    for i, palavra in enumerate(palavras):
        if i > 0 and palavra in preposicoes:
            palavras[i] = palavra.lower()
    return ' '.join(palavras)


def _formatar_email(valor):
    """Email sempre em minúsculo, sem espaços."""
    if pd.isna(valor) or str(valor).strip() == '':
        return ''
    return str(valor).strip().lower()


def _formatar_estado(valor):
    """Estado sempre com 2 caracteres maiúsculos."""
    if pd.isna(valor) or str(valor).strip() == '':
        return ''
    estado = str(valor).strip().upper()
    # Se já tem 2 caracteres, retorna direto
    if len(estado) == 2:
        return estado
    # Mapa de nomes completos para siglas
    estados_map = {
        'ACRE': 'AC', 'ALAGOAS': 'AL', 'AMAPÁ': 'AP', 'AMAZONAS': 'AM',
        'BAHIA': 'BA', 'CEARÁ': 'CE', 'DISTRITO FEDERAL': 'DF', 'ESPÍRITO SANTO': 'ES',
        'GOIÁS': 'GO', 'MARANHÃO': 'MA', 'MATO GROSSO': 'MT', 'MATO GROSSO DO SUL': 'MS',
        'MINAS GERAIS': 'MG', 'PARÁ': 'PA', 'PARAÍBA': 'PB', 'PARANÁ': 'PR',
        'PERNAMBUCO': 'PE', 'PIAUÍ': 'PI', 'RIO DE JANEIRO': 'RJ',
        'RIO GRANDE DO NORTE': 'RN', 'RIO GRANDE DO SUL': 'RS', 'RONDÔNIA': 'RO',
        'RORAIMA': 'RR', 'SANTA CATARINA': 'SC', 'SÃO PAULO': 'SP',
        'SERGIPE': 'SE', 'TOCANTINS': 'TO',
    }
    return estados_map.get(estado, estado[:2])


def _aplicar_formatacoes(df: pd.DataFrame, tipo: str) -> pd.DataFrame:
    """Aplica formatações específicas por tipo de dado."""
    df = df.copy()
    
    # === REGRAS GLOBAIS (todos os tipos) ===
    
    # Email: sempre minúsculo
    campos_email = ['EMAIL', 'E-MAIL']
    for col in campos_email:
        if col in df.columns:
            df[col] = df[col].apply(_formatar_email)
    
    # Estado: sempre 2 caracteres maiúsculos
    campos_estado = ['ESTADO 1', 'ESTADO 2', 'UF']
    for col in campos_estado:
        if col in df.columns:
            df[col] = df[col].apply(_formatar_estado)
    
    # Telefones: padrão +55 (XX) XXXXX-XXXX
    campos_telefone = ['TELEFONE COMERCIAL', 'TELEFONE RESIDENCIAL', 'CELULAR', 'FAX']
    for col in campos_telefone:
        if col in df.columns:
            df[col] = df[col].apply(_formatar_telefone)
    
    # === REGRAS ESPECÍFICAS POR TIPO ===
    
    if tipo == 'contatos':
        # Formatar documentos
        if 'DOCUMENTO FEDERAL' in df.columns:
            df['DOCUMENTO FEDERAL'] = df['DOCUMENTO FEDERAL'].apply(_formatar_cpf_cnpj)
        
        # Formatar CEP
        for col in ['CEP 1', 'CEP 2']:
            if col in df.columns:
                df[col] = df[col].apply(_formatar_cep)
    
    elif tipo == 'financeiro':
        # Formatar datas
        for col in ['DATA', 'EMISSÃO', 'VENCIMENTO']:
            if col in df.columns:
                df[col] = df[col].apply(_converter_data)
        
        # Formatar valor como numérico
        if 'VALOR' in df.columns:
            def _converter_valor(val):
                if pd.isna(val) or str(val).strip() == '':
                    return None
                # Se já é numérico, retorna direto
                if isinstance(val, (int, float)):
                    return float(val)
                val_str = str(val).strip().replace('R$', '').replace(' ', '')
                # Se tem ponto E vírgula, ponto é milhar (1.500,00)
                if '.' in val_str and ',' in val_str:
                    val_str = val_str.replace('.', '').replace(',', '.')
                # Se tem só vírgula, é decimal (1500,00)
                elif ',' in val_str:
                    val_str = val_str.replace(',', '.')
                # Se tem só ponto, é decimal (padrão internacional)
                try:
                    return float(val_str)
                except ValueError:
                    return None
            df['VALOR'] = df['VALOR'].apply(_converter_valor)
        
        # Traduzir TIPO DE RECEITA/DESPESA para chave DOit
        if 'TIPO DE RECEITA/DESPESA' in df.columns:
            df['TIPO DE RECEITA/DESPESA'] = df['TIPO DE RECEITA/DESPESA'].apply(_traduzir_tipo_receita_despesa)
    
    elif tipo == 'horas':
        # Formatar datas
        for col in ['DATA INICIAL', 'DATA FINAL']:
            if col in df.columns:
                df[col] = df[col].apply(_converter_data)
        
        # Converter horas para numérico
        if 'HORAS TRABALHADAS' in df.columns:
            df['HORAS TRABALHADAS'] = pd.to_numeric(
                df['HORAS TRABALHADAS'].astype(str).str.replace(',', '.', regex=False),
                errors='coerce'
            )
    
    elif tipo == 'projetos':
        # Formatar datas
        for col in ['INÍCIO', 'EXECUÇÃO', 'TÉRMINO']:
            if col in df.columns:
                df[col] = df[col].apply(_converter_data)
    
    elif tipo == 'usuarios':
        # LOGIN: primeiro nome em minúsculo
        if 'LOGIN' in df.columns and 'NOME' in df.columns:
            df['LOGIN'] = df['NOME'].apply(
                lambda x: str(x).strip().split()[0].lower() if pd.notna(x) and str(x).strip() else ''
            )
        
        # SENHA: baseada no CARGO/FUNÇÃO
        if 'SENHA' in df.columns:
            def _definir_senha(row):
                cargo = str(row.get('_cargo_original', '')).strip().lower() if pd.notna(row.get('_cargo_original')) else ''
                funcao = str(row.get('_funcao_original', '')).strip().lower() if pd.notna(row.get('_funcao_original')) else ''
                texto = cargo + ' ' + funcao
                # Perfis administrativos/financeiros/sócios → 789
                if any(p in texto for p in ['admin', 'financ', 'sócio', 'socio', 'diretor', 'gerente', 'gestor', 'coordenador']):
                    return '789'
                # Demais (arquitetos, estagiários, designers, etc) → 123
                return '123'
            df['SENHA'] = df.apply(_definir_senha, axis=1)
        
        # ATIVO: todos ativos
        if 'ATIVO' in df.columns:
            df['ATIVO'] = 'Sim'
        
        # EMAIL: configurar HOST, PORTA, SSL, TLS, USUÁRIO (SMTP) com base no servidor
        if 'EMAIL' in df.columns:
            df['EMAIL'] = df['EMAIL'].apply(lambda x: str(x).strip().lower() if pd.notna(x) and str(x).strip() else '')
            
            def _configurar_email(row):
                email = str(row.get('EMAIL', '')).strip().lower()
                servidor = str(row.get('_servidor_original', '')).strip().lower() if pd.notna(row.get('_servidor_original')) else ''
                
                # Detectar provedor pelo email ou campo servidor
                host = ''
                porta = ''
                ssl = ''
                tls = ''
                
                if 'gmail' in email or 'gmail' in servidor:
                    host = 'smtp.gmail.com'
                    porta = '587'
                    ssl = 'Não'
                    tls = 'Sim'
                elif 'outlook' in email or 'hotmail' in email or 'outlook' in servidor or 'microsoft' in servidor:
                    host = 'smtp.office365.com'
                    porta = '587'
                    ssl = 'Não'
                    tls = 'Sim'
                elif 'yahoo' in email or 'yahoo' in servidor:
                    host = 'smtp.mail.yahoo.com'
                    porta = '465'
                    ssl = 'Sim'
                    tls = 'Não'
                elif email:
                    # Domínio próprio: smtp.dominio.com
                    dominio = email.split('@')[1] if '@' in email else ''
                    if dominio:
                        host = f'smtp.{dominio}'
                        porta = '587'
                        ssl = 'Não'
                        tls = 'Sim'
                
                row['HOST'] = host
                row['PORTA'] = porta
                row['USAR SSL'] = ssl
                row['USAR TLS'] = tls
                row['USUÁRIO (SMTP)'] = email
                return row
            
            df = df.apply(_configurar_email, axis=1)
        
        # PERMISSÕES: baseadas na FUNÇÃO
        if 'BASICO?' in df.columns:
            def _configurar_permissoes(row):
                funcao = str(row.get('_funcao_original', '')).strip().lower() if pd.notna(row.get('_funcao_original')) else ''
                cargo = str(row.get('_cargo_original', '')).strip().lower() if pd.notna(row.get('_cargo_original')) else ''
                texto = funcao + ' ' + cargo
                
                # Todos têm básico e projeto 1
                row['BASICO?'] = 'Sim'
                row['PROJ. 1?'] = 'Sim'
                
                # Líderes: tudo + projeto 2
                if any(p in texto for p in ['líder', 'lider', 'coordenador', 'supervisor']):
                    row['PROJ. 2?'] = 'Sim'
                    row['FIN.?'] = 'Sim'
                    row['GERENTE?'] = 'Sim'
                    row['COMPRAS?'] = 'Sim'
                    row['FATUR.?'] = 'Sim'
                    row['VENDAS?'] = 'Sim'
                
                # Sócios e financeiro: tudo + gerente + projeto 3 + financeiro
                if any(p in texto for p in ['sócio', 'socio', 'diretor', 'admin', 'financ', 'gerente', 'gestor']):
                    row['PROJ. 2?'] = 'Sim'
                    row['PROJ. 3?'] = 'Sim'
                    row['FIN.?'] = 'Sim'
                    row['GERENTE?'] = 'Sim'
                    row['COMPRAS?'] = 'Sim'
                    row['FATUR.?'] = 'Sim'
                    row['VENDAS?'] = 'Sim'
                    row['ADMIN?'] = 'Sim'
                
                return row
            
            df = df.apply(_configurar_permissoes, axis=1)
    
    return df


def _aplicar_caixa(df: pd.DataFrame, estilo_caixa: str) -> pd.DataFrame:
    """
    Aplica estilo de caixa (MAIÚSCULA, Primeira Maiúscula ou ORIGINAL) nos campos de texto.
    Não afeta email (sempre minúsculo) nem estado (sempre 2 chars maiúsculo).
    Se ORIGINAL, não altera nada.
    """
    if estilo_caixa == 'ORIGINAL':
        return df
    
    df = df.copy()
    campos_texto = ['NOME', 'APELIDO', 'RUA 1', 'RUA 2', 'BAIRRO 1', 'BAIRRO 2',
                    'CIDADE 1', 'CIDADE 2', 'COMPLEMENTO 1', 'COMPLEMENTO 2',
                    'PAÍS 1', 'PAÍS 2', 'CLASSIFICAÇÃO', 'GERENTE DE CONTA',
                    'DESCRIÇÃO', 'NOME DA ATIVIDADE', 'DESCRIÇÃO DA ATIVIDADE',
                    'DEPARTAMENTO', '1ª CATEGORIA', '2ª CATEGORIA', '3ª CATEGORIA',
                    'CATEGORIA', 'ETAPA 1', 'ETAPA 2', 'ETAPA 3', 'ANOTAÇÕES',
                    'TIPO DE ENDEREÇO 1', 'TIPO DE ENDEREÇO 2', 'ORIGEM',
                    'INFORMAÇÕES BANCÁRIAS', 'FORMA DE PAGAMENTO', 'TIPO',
                    'STATUS']
    
    for col in campos_texto:
        if col in df.columns:
            if estilo_caixa == 'MAIÚSCULA':
                df[col] = df[col].apply(
                    lambda x: str(x).strip().upper() if pd.notna(x) and str(x).strip() else ''
                )
            else:  # Primeira Maiúscula
                df[col] = df[col].apply(_primeira_maiuscula)
    
    return df


def _campos_extras_para_anotacoes(df_entrada: pd.DataFrame, mapeamento: dict, colunas_padrao: list) -> pd.Series:
    """
    Pega colunas da origem que NÃO foram mapeadas para nenhum campo padrão
    e junta tudo no campo ANOTAÇÕES.
    Formato: "Campo: Valor | Campo2: Valor2"
    Ignora valores vazios, FALSE, e campos que já foram mapeados.
    """
    # Todas as colunas de origem que foram usadas em algum mapeamento
    usadas = set(v for v in mapeamento.values() if v is not None)
    colunas_extras = [c for c in df_entrada.columns if c not in usadas]
    
    if not colunas_extras:
        return pd.Series([''] * len(df_entrada))
    
    def _montar_anotacao(row):
        partes = []
        for col in colunas_extras:
            val = row.get(col)
            if pd.notna(val) and str(val).strip() and str(val).strip().upper() not in ('FALSE', 'NAN', 'NONE', ''):
                partes.append(f"{col}: {val}")
        return ' | '.join(partes)
    
    return df_entrada.apply(_montar_anotacao, axis=1)


def _detectar_tipo_coluna(serie: pd.Series) -> str:
    """
    Detecta o tipo de dado de uma coluna para distribuir em colunas customizáveis.
    Retorna: 'booleana', 'moeda', 'data', 'numero', 'texto'
    """
    valores = serie.dropna()
    if valores.empty:
        return 'texto'
    
    valores_str = valores.astype(str).str.strip()
    valores_str = valores_str[valores_str != '']
    if valores_str.empty:
        return 'texto'
    
    # Verificar booleano
    valores_lower = valores_str.str.lower()
    valores_bool = {'true', 'false', 'sim', 'não', 'nao', 'yes', 'no', '0', '1', 'verdadeiro', 'falso'}
    if valores_lower.isin(valores_bool).mean() > 0.8:
        return 'booleana'
    
    # Verificar número puro (inteiros ou decimais simples, sem formato de data)
    numeros = pd.to_numeric(valores_str.str.replace(',', '.', regex=False), errors='coerce')
    if numeros.notna().mean() > 0.8:
        # Verificar se parece moeda (valores grandes com decimais)
        if numeros.mean() > 100 and (valores_str.str.contains(r'\d{4,}', regex=True).mean() > 0.5):
            return 'moeda'
        return 'numero'
    
    # Verificar moeda (R$ 1.500,00 ou 1500.00)
    padrao_moeda = valores_str.str.match(r'^[R$\s]*[\d.]+[,]\d{2}$|^[R$\s]*[\d,]+[.]\d{2}$')
    if padrao_moeda.mean() > 0.5:
        return 'moeda'
    
    # Verificar data (formato dd/mm/yyyy, yyyy-mm-dd, etc)
    # Só considerar data se tiver separadores típicos de data
    tem_formato_data = valores_str.str.match(r'^\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4}$|^\d{4}[/\-\.]\d{1,2}[/\-\.]\d{1,2}')
    if tem_formato_data.mean() > 0.5:
        return 'data'
    
    return 'texto'


def _distribuir_extras_em_customizaveis(df_entrada: pd.DataFrame, df_saida: pd.DataFrame, mapeamento: dict) -> pd.DataFrame:
    """
    Para projetos: distribui campos extras nas colunas customizáveis
    de acordo com o tipo de dado detectado.
    Retorna o df_saida atualizado + dicionário de mapeamento customizável.
    """
    usadas = set(v for v in mapeamento.values() if v is not None)
    colunas_extras = [c for c in df_entrada.columns if c not in usadas]
    
    if not colunas_extras:
        return df_saida, {}
    
    # Contadores por tipo
    slots = {
        'booleana': ['COLUNA CUSTOMIZÁVEL BOOLEANA 1', 'COLUNA CUSTOMIZÁVEL BOOLEANA 2', 'COLUNA CUSTOMIZÁVEL BOOLEANA 3'],
        'moeda': ['COLUNA CUSTOMIZÁVEL MOEDA 1', 'COLUNA CUSTOMIZÁVEL MOEDA 2', 'COLUNA CUSTOMIZÁVEL MOEDA 3'],
        'data': ['COLUNA CUSTOMIZÁVEL DATA 1', 'COLUNA CUSTOMIZÁVEL DATA 2', 'COLUNA CUSTOMIZÁVEL DATA 3',
                 'COLUNA CUSTOMIZÁVEL DATA 4', 'COLUNA CUSTOMIZÁVEL DATA 5', 'COLUNA CUSTOMIZÁVEL DATA 6'],
        'numero': ['COLUNA CUSTOMIZÁVEL NUMERO 1', 'COLUNA CUSTOMIZÁVEL NUMERO 2', 'COLUNA CUSTOMIZÁVEL NUMERO 3'],
        'texto': ['COLUNA CUSTOMIZÁVEL TEXTO 1', 'COLUNA CUSTOMIZÁVEL TEXTO 2', 'COLUNA CUSTOMIZÁVEL TEXTO 3',
                  'COLUNA CUSTOMIZÁVEL TEXTO 4', 'COLUNA CUSTOMIZÁVEL TEXTO 5'],
    }
    
    idx_usado = {tipo: 0 for tipo in slots}
    mapeamento_custom = {}  # {coluna_customizavel: nome_original_do_campo}
    extras_sem_slot = []  # Campos que não couberam
    
    for col_extra in colunas_extras:
        # Verificar se a coluna tem dados
        if df_entrada[col_extra].dropna().empty:
            continue
        if df_entrada[col_extra].astype(str).str.strip().replace('', pd.NA).dropna().empty:
            continue
            
        tipo = _detectar_tipo_coluna(df_entrada[col_extra])
        
        if idx_usado[tipo] < len(slots[tipo]):
            col_destino = slots[tipo][idx_usado[tipo]]
            if col_destino in df_saida.columns:
                df_saida[col_destino] = df_entrada[col_extra].values
                mapeamento_custom[col_destino] = col_extra
                idx_usado[tipo] += 1
            else:
                extras_sem_slot.append(col_extra)
        else:
            # Tentar colocar em texto se não couber no tipo original
            if tipo != 'texto' and idx_usado['texto'] < len(slots['texto']):
                col_destino = slots['texto'][idx_usado['texto']]
                if col_destino in df_saida.columns:
                    df_saida[col_destino] = df_entrada[col_extra].astype(str).values
                    mapeamento_custom[col_destino] = col_extra
                    idx_usado['texto'] += 1
                else:
                    extras_sem_slot.append(col_extra)
            else:
                extras_sem_slot.append(col_extra)
    
    return df_saida, mapeamento_custom


def _carregar_padronizacoes():
    """
    Carrega listas de padronização do DOit (classificações, categorias, etc).
    Retorna dicionário com as listas válidas.
    """
    padronizacoes = {}
    arquivos_padrao = {
        'classificacoes_cadastro': 'classificacoes_cadastro',
        'categorias_projeto': 'categorias_projeto',
        'categorias_financeiro': 'categorias_financeiro',
        'departamentos_financeiro': 'departamentos_financeiro',
        'formas_pagamento': 'formas_pagamento',
        'tipos_endereco': 'tipos_endereco',
    }
    
    MAPPINGS_DIR = os.path.join(BASE_DIR, '..', 'mappings')
    
    for chave, nome_base in arquivos_padrao.items():
        for ext in ['.xlsx', '.csv', '.txt']:
            # Procurar em models e mappings
            for pasta in [MODELOS_DIR, MAPPINGS_DIR]:
                caminho = os.path.join(pasta, f"{nome_base}{ext}")
                if os.path.exists(caminho):
                    try:
                        if ext == '.xlsx':
                            df = pd.read_excel(caminho)
                            padronizacoes[chave] = df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
                        elif ext == '.csv':
                            df = pd.read_csv(caminho)
                            padronizacoes[chave] = df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
                        else:  # .txt
                            with open(caminho, 'r', encoding='utf-8') as f:
                                padronizacoes[chave] = [l.strip() for l in f.readlines() if l.strip()]
                    except Exception:
                        pass
                    break
            if chave in padronizacoes:
                break
    
    return padronizacoes


def _validar_padronizacoes(df: pd.DataFrame, tipo: str, padronizacoes: dict) -> list:
    """
    Valida se os valores estão dentro das padronizações do DOit.
    Retorna lista de alertas sobre valores novos que precisam ser criados.
    """
    alertas = []
    
    if tipo == 'contatos':
        # Validar classificações
        if 'classificacoes_cadastro' in padronizacoes and 'CLASSIFICAÇÃO' in df.columns:
            valores_validos = [v.lower() for v in padronizacoes['classificacoes_cadastro']]
            valores_usados = df['CLASSIFICAÇÃO'].dropna().unique()
            for val in valores_usados:
                val_str = str(val).strip()
                if val_str and val_str.lower() not in valores_validos:
                    alertas.append(f"🏷️ Classificação de cadastro nova: **{val_str}** (não existe no DOit)")
        
        # Validar tipos de endereço
        if 'tipos_endereco' in padronizacoes:
            for col in ['TIPO DE ENDEREÇO 1', 'TIPO DE ENDEREÇO 2']:
                if col in df.columns:
                    valores_validos = [v.lower() for v in padronizacoes['tipos_endereco']]
                    valores_usados = df[col].dropna().unique()
                    for val in valores_usados:
                        val_str = str(val).strip()
                        if val_str and val_str.lower() not in valores_validos:
                            alertas.append(f"📍 Tipo de endereço novo em {col}: **{val_str}**")
        
        # Validar forma de pagamento
        if 'formas_pagamento' in padronizacoes and 'FORMA DE PAGAMENTO' in df.columns:
            valores_validos = [v.lower() for v in padronizacoes['formas_pagamento']]
            valores_usados = df['FORMA DE PAGAMENTO'].dropna().unique()
            for val in valores_usados:
                val_str = str(val).strip()
                if val_str and val_str.lower() not in valores_validos:
                    alertas.append(f"💳 Forma de pagamento nova: **{val_str}**")
    
    if tipo == 'projetos':
        if 'categorias_projeto' in padronizacoes and 'CATEGORIA' in df.columns:
            valores_validos = [v.lower() for v in padronizacoes['categorias_projeto']]
            valores_usados = df['CATEGORIA'].dropna().unique()
            for val in valores_usados:
                val_str = str(val).strip()
                if val_str and val_str.lower() not in valores_validos:
                    alertas.append(f"📂 Categoria de projeto nova: **{val_str}** (não existe no DOit)")
    
    if tipo == 'financeiro':
        # Validar departamento
        if 'departamentos_financeiro' in padronizacoes and 'DEPARTAMENTO' in df.columns:
            valores_validos = [v.lower() for v in padronizacoes['departamentos_financeiro']]
            valores_usados = df['DEPARTAMENTO'].dropna().unique()
            for val in valores_usados:
                val_str = str(val).strip()
                if val_str and val_str.lower() not in valores_validos:
                    alertas.append(f"🏢 Departamento novo: **{val_str}** (não existe no DOit)")
        
        # Validar categorias (plano de contas - 3 níveis)
        if 'categorias_financeiro' in padronizacoes:
            # Extrair valores válidos por nível
            nivel1_validos = set()
            nivel2_validos = set()
            nivel3_validos = set()
            for linha in padronizacoes['categorias_financeiro']:
                partes = [p.strip() for p in linha.split('>')]
                if len(partes) >= 1 and partes[0]:
                    nivel1_validos.add(partes[0].lower())
                if len(partes) >= 2 and partes[1]:
                    nivel2_validos.add(partes[1].lower())
                if len(partes) >= 3 and partes[2]:
                    nivel3_validos.add(partes[2].lower())
            
            # Validar 1ª CATEGORIA (Nível 1)
            if '1ª CATEGORIA' in df.columns:
                valores_usados = df['1ª CATEGORIA'].dropna().unique()
                for val in valores_usados:
                    val_str = str(val).strip()
                    if val_str and val_str.lower() not in nivel1_validos:
                        alertas.append(f"📊 1ª Categoria nova: **{val_str}** (não existe no plano de contas)")
            
            # Validar 2ª CATEGORIA (Nível 2)
            if '2ª CATEGORIA' in df.columns:
                valores_usados = df['2ª CATEGORIA'].dropna().unique()
                for val in valores_usados:
                    val_str = str(val).strip()
                    if val_str and val_str.lower() not in nivel2_validos:
                        alertas.append(f"📊 2ª Categoria nova: **{val_str}** (não existe no plano de contas)")
            
            # Validar 3ª CATEGORIA (Nível 3)
            if '3ª CATEGORIA' in df.columns:
                valores_usados = df['3ª CATEGORIA'].dropna().unique()
                for val in valores_usados:
                    val_str = str(val).strip()
                    if val_str and val_str.lower() not in nivel3_validos:
                        alertas.append(f"📊 3ª Categoria nova: **{val_str}** (não existe no plano de contas)")
        
        # Validar forma de pagamento (se existir no financeiro)
        if 'formas_pagamento' in padronizacoes and 'FORMA DE PAGAMENTO' in df.columns:
            valores_validos = [v.lower() for v in padronizacoes['formas_pagamento']]
            valores_usados = df['FORMA DE PAGAMENTO'].dropna().unique()
            for val in valores_usados:
                val_str = str(val).strip()
                if val_str and val_str.lower() not in valores_validos:
                    alertas.append(f"💳 Forma de pagamento nova: **{val_str}**")
        
        # Validar TIPO DE RECEITA/DESPESA
        if 'TIPO DE RECEITA/DESPESA' in df.columns:
            chaves_validas = set(TIPOS_RECEITA_DESPESA.keys())
            valores_usados = df['TIPO DE RECEITA/DESPESA'].dropna().unique()
            for val in valores_usados:
                val_str = str(val).strip()
                if val_str and val_str not in chaves_validas:
                    alertas.append(f"📑 Tipo de Receita/Despesa não reconhecido: **{val_str}** (não foi possível traduzir para chave DOit)")
    
    return alertas


def converter_arquivo(
    arquivo: str,
    tipo: str,
    origem: str = 'excel_manual',
    mapeamento_custom: dict = None,
    id_inicial: int = 15,
    estilo_caixa: str = 'Primeira Maiúscula',
    df_cadastro_ref: pd.DataFrame = None,
    df_projetos_ref: pd.DataFrame = None,
) -> dict:
    """
    Converte um arquivo de entrada para o padrão DOit.
    
    Args:
        arquivo: Caminho do arquivo de entrada (.xlsx, .xls, .csv)
        tipo: Tipo de dado ('contatos', 'projetos', 'financeiro', 'horas')
        origem: Sistema de origem ('clickup', 'sienge', 'trello', 'excel_manual')
        mapeamento_custom: Dicionário customizado {coluna_padrao: coluna_origem}
        id_inicial: ID inicial para sequencial (default: 15)
        estilo_caixa: 'Primeira Maiúscula' ou 'MAIÚSCULA'
        df_cadastro_ref: DataFrame de cadastro já convertido (para vincular IDs em projetos/financeiro)
        df_projetos_ref: DataFrame de projetos já convertido (para vincular IDs no financeiro)
    
    Returns:
        dict com:
            'dados': DataFrame no formato padrão DOit
            'alertas': lista de alertas sobre padronizações
    """
    # Carregar arquivo de entrada
    if arquivo.endswith('.csv'):
        df_entrada = pd.read_csv(arquivo)
    else:
        df_entrada = pd.read_excel(arquivo)
    
    if df_entrada.empty:
        raise ValueError("Arquivo de entrada está vazio.")
    
    # === LIMPAR LINHAS DE DESCRIÇÃO/EXEMPLO (DOit Coleta) ===
    if origem == 'doit_coleta':
        # A planilha de coleta tem:
        # Linha 1 = cabeçalhos (já lida pelo pandas como header)
        # Linhas seguintes podem ter DESCRIÇÃO e EXEMPLO que devem ser removidas
        # Detectar pela primeira coluna (ou qualquer coluna) começando com "DESCRIÇÃO" ou "EXEMPLO"
        primeira_col = df_entrada.iloc[:, 0] if len(df_entrada.columns) > 0 else pd.Series()
        if not primeira_col.empty:
            mask_remover = primeira_col.astype(str).str.strip().str.upper().str.startswith(('DESCRIÇÃO', 'EXEMPLO'))
            df_entrada = df_entrada[~mask_remover].reset_index(drop=True)
        
        # Remover linhas completamente vazias
        df_entrada = df_entrada.dropna(how='all').reset_index(drop=True)
    
    if df_entrada.empty:
        raise ValueError("Arquivo de entrada está vazio após remover linhas de descrição/exemplo.")
    
    # Carregar colunas do modelo padrão
    colunas_padrao = carregar_modelo(tipo)
    
    # Determinar mapeamento
    if mapeamento_custom:
        mapeamento = mapeamento_custom
    elif origem in MAPEAMENTOS and tipo in MAPEAMENTOS[origem]:
        mapeamento_definido = MAPEAMENTOS[origem][tipo]
        mapeamento = {}
        for col_padrao, opcoes in mapeamento_definido.items():
            col_encontrada = _encontrar_coluna(list(df_entrada.columns), opcoes)
            if col_encontrada:
                mapeamento[col_padrao] = col_encontrada
        
        # Complementar com mapeamento automático para colunas não mapeadas
        auto = _mapear_automatico(df_entrada, [c for c in colunas_padrao if c not in mapeamento])
        mapeamento.update(auto)
    else:
        mapeamento = _mapear_automatico(df_entrada, colunas_padrao)
    
    # Construir DataFrame de saída
    df_saida = pd.DataFrame(columns=colunas_padrao)
    
    for col_padrao in colunas_padrao:
        if col_padrao in mapeamento and mapeamento[col_padrao] in df_entrada.columns:
            df_saida[col_padrao] = df_entrada[mapeamento[col_padrao]].values
        else:
            df_saida[col_padrao] = ''
    
    # === ID SEQUENCIAL ===
    if 'ID' in df_saida.columns:
        df_saida['ID'] = range(id_inicial, id_inicial + len(df_saida))
    
    # === CAMPOS EXTRAS ===
    mapeamento_custom = {}
    if tipo == 'contatos' and 'ANOTAÇÕES' in df_saida.columns:
        # Contatos: extras vão para ANOTAÇÕES
        anotacoes_extras = _campos_extras_para_anotacoes(df_entrada, mapeamento, colunas_padrao)
        existente = df_saida['ANOTAÇÕES'].apply(lambda x: str(x).strip() if pd.notna(x) and str(x).strip() else '')
        df_saida['ANOTAÇÕES'] = existente.combine(anotacoes_extras, 
            lambda a, b: f"{a} | {b}" if a and b else (a or b))
    
    elif tipo == 'projetos':
        # Projetos: extras vão para colunas customizáveis por tipo de dado
        df_saida, mapeamento_custom = _distribuir_extras_em_customizaveis(df_entrada, df_saida, mapeamento)
    
    elif tipo in ('financeiro', 'horas'):
        # Financeiro/Horas: extras vão para anotações/descrição se existir
        col_anotacao = 'DESCRIÇÃO' if 'DESCRIÇÃO' in df_saida.columns else None
        if col_anotacao:
            usadas = set(v for v in mapeamento.values() if v is not None)
            colunas_extras = [c for c in df_entrada.columns if c not in usadas]
            if colunas_extras:
                def _montar_extra(row):
                    partes = []
                    for col in colunas_extras:
                        val = row.get(col)
                        if pd.notna(val) and str(val).strip() and str(val).strip().upper() not in ('FALSE', 'NAN', 'NONE', ''):
                            partes.append(f"{col}: {val}")
                    return ' | '.join(partes)
                extras = df_entrada.apply(_montar_extra, axis=1)
                existente = df_saida[col_anotacao].apply(lambda x: str(x).strip() if pd.notna(x) and str(x).strip() else '')
                df_saida[col_anotacao] = existente.combine(extras,
                    lambda a, b: f"{a} | {b}" if a and b else (a or b))
    
    # === VINCULAR ID DO CADASTRO (para projetos) ===
    if tipo == 'projetos' and df_cadastro_ref is not None and 'ID DO CADASTRO' in df_saida.columns:
        # Tentar vincular pelo nome do cliente
        if 'NOME' in df_cadastro_ref.columns:
            mapa_nome_id = dict(zip(
                df_cadastro_ref['NOME'].str.lower().str.strip(),
                df_cadastro_ref['ID']
            ))
            # Procurar coluna de cliente na entrada
            col_cliente_origem = _encontrar_coluna(
                list(df_entrada.columns),
                ['Cliente', 'cliente', 'Card Name', 'Nome do Cliente', 'Razão Social', 'Nome']
            )
            if col_cliente_origem:
                df_saida['ID DO CADASTRO'] = df_entrada[col_cliente_origem].apply(
                    lambda x: mapa_nome_id.get(str(x).lower().strip(), '') if pd.notna(x) else ''
                )
    
    # === VINCULAR IDs (para financeiro) ===
    if tipo == 'financeiro':
        if df_cadastro_ref is not None and 'ID DE / PARA' in df_saida.columns:
            if 'NOME' in df_cadastro_ref.columns:
                mapa_nome_id = dict(zip(
                    df_cadastro_ref['NOME'].str.lower().str.strip(),
                    df_cadastro_ref['ID']
                ))
                col_cliente_origem = _encontrar_coluna(
                    list(df_entrada.columns),
                    ['Cliente', 'cliente', 'Fornecedor', 'Nome', 'Razão Social', 'De/Para']
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
                col_projeto_origem = _encontrar_coluna(
                    list(df_entrada.columns),
                    ['Projeto', 'projeto', 'Obra', 'Nome do Projeto', 'Project']
                )
                if col_projeto_origem:
                    df_saida['ID PROJETO'] = df_entrada[col_projeto_origem].apply(
                        lambda x: mapa_proj_id.get(str(x).lower().strip(), '') if pd.notna(x) else ''
                    )
    
    # === APLICAR FORMATAÇÕES BASE ===
    df_saida = _aplicar_formatacoes(df_saida, tipo)
    
    # === APLICAR ESTILO DE CAIXA (sobrescreve a formatação de texto) ===
    df_saida = _aplicar_caixa(df_saida, estilo_caixa)
    
    # === VALIDAR PADRONIZAÇÕES ===
    padronizacoes = _carregar_padronizacoes()
    alertas = _validar_padronizacoes(df_saida, tipo, padronizacoes)
    
    # Adicionar info sobre colunas customizáveis usadas
    if tipo == 'projetos' and mapeamento_custom:
        alertas.insert(0, "📋 **Colunas customizáveis utilizadas (renomear no DOit depois):**")
        for col_custom, col_original in mapeamento_custom.items():
            alertas.insert(1, f"   ↳ {col_custom} ← {col_original}")
    
    return {
        'dados': df_saida,
        'alertas': alertas,
    }


def obter_preview_mapeamento(
    arquivo: str,
    tipo: str,
    origem: str = 'excel_manual'
) -> dict:
    """
    Retorna um preview do mapeamento que será aplicado, sem converter.
    Útil para a interface Streamlit mostrar ao usuário antes de confirmar.
    
    Returns:
        {
            'colunas_padrao': [...],
            'colunas_origem': [...],
            'mapeamento': {col_padrao: col_origem ou None},
            'nao_mapeadas_origem': [...],  # colunas da origem sem correspondência
        }
    """
    if arquivo.endswith('.csv'):
        df_entrada = pd.read_csv(arquivo, nrows=5)
    else:
        df_entrada = pd.read_excel(arquivo, nrows=5)
    
    colunas_padrao = carregar_modelo(tipo)
    colunas_origem = list(df_entrada.columns)
    
    # Determinar mapeamento
    if origem in MAPEAMENTOS and tipo in MAPEAMENTOS[origem]:
        mapeamento_definido = MAPEAMENTOS[origem][tipo]
        mapeamento = {}
        for col_padrao, opcoes in mapeamento_definido.items():
            col_encontrada = _encontrar_coluna(colunas_origem, opcoes)
            mapeamento[col_padrao] = col_encontrada
        
        # Complementar com automático
        auto = _mapear_automatico(df_entrada, [c for c in colunas_padrao if c not in mapeamento])
        for k, v in auto.items():
            if k not in mapeamento:
                mapeamento[k] = v
    else:
        mapeamento = _mapear_automatico(df_entrada, colunas_padrao)
    
    # Preencher None para não mapeadas
    for col in colunas_padrao:
        if col not in mapeamento:
            mapeamento[col] = None
    
    # Colunas da origem que não foram usadas
    usadas = set(v for v in mapeamento.values() if v is not None)
    nao_mapeadas = [c for c in colunas_origem if c not in usadas]
    
    return {
        'colunas_padrao': colunas_padrao,
        'colunas_origem': colunas_origem,
        'mapeamento': mapeamento,
        'nao_mapeadas_origem': nao_mapeadas,
    }


# ============================================================
# CLI
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description='Conversor de dados para padrão DOit'
    )
    parser.add_argument('--arquivo', '-a', required=True, help='Arquivo de entrada (.xlsx, .csv)')
    parser.add_argument('--tipo', '-t', required=True, 
                       choices=['contatos', 'projetos', 'financeiro', 'horas'],
                       help='Tipo de dado')
    parser.add_argument('--origem', '-o', default='excel_manual',
                       choices=['clickup', 'sienge', 'trello', 'excel_manual'],
                       help='Sistema de origem (default: excel_manual)')
    parser.add_argument('--saida', '-s', help='Arquivo de saída (default: <arquivo>_convertido.xlsx)')
    parser.add_argument('--id-inicial', '-i', type=int, default=15,
                       help='ID inicial para sequencial (default: 15)')
    parser.add_argument('--caixa', '-c', choices=['primeira', 'maiuscula'], default='primeira',
                       help='Estilo de caixa: "primeira" (Primeira Maiúscula) ou "maiuscula" (TUDO MAIÚSCULO)')
    parser.add_argument('--cadastro-ref', help='Arquivo de cadastro já convertido (para vincular IDs)')
    parser.add_argument('--projetos-ref', help='Arquivo de projetos já convertido (para vincular IDs)')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.arquivo):
        print(f"❌ Arquivo não encontrado: {args.arquivo}")
        sys.exit(1)
    
    # Definir nome de saída
    if args.saida:
        saida = args.saida
    else:
        nome_base = os.path.splitext(os.path.basename(args.arquivo))[0]
        saida = f"{nome_base}_convertido.xlsx"
    
    estilo_caixa = 'MAIÚSCULA' if args.caixa == 'maiuscula' else 'Primeira Maiúscula'
    
    print(f"📂 Arquivo: {args.arquivo}")
    print(f"📋 Tipo: {args.tipo}")
    print(f"🔄 Origem: {args.origem}")
    print(f"🔢 ID inicial: {args.id_inicial}")
    print(f"🔤 Caixa: {estilo_caixa}")
    print(f"💾 Saída: {saida}")
    print()
    
    # Carregar referências se fornecidas
    df_cadastro_ref = None
    df_projetos_ref = None
    if args.cadastro_ref and os.path.exists(args.cadastro_ref):
        df_cadastro_ref = pd.read_excel(args.cadastro_ref)
        print(f"📎 Cadastro referência: {args.cadastro_ref} ({len(df_cadastro_ref)} registros)")
    if args.projetos_ref and os.path.exists(args.projetos_ref):
        df_projetos_ref = pd.read_excel(args.projetos_ref)
        print(f"📎 Projetos referência: {args.projetos_ref} ({len(df_projetos_ref)} registros)")
    
    # Preview do mapeamento
    preview = obter_preview_mapeamento(args.arquivo, args.tipo, args.origem)
    print("\n📊 Mapeamento de colunas:")
    for col_padrao, col_origem in preview['mapeamento'].items():
        status = f"← {col_origem}" if col_origem else "⚠️  (vazio)"
        print(f"   {col_padrao:30s} {status}")
    
    if preview['nao_mapeadas_origem']:
        print(f"\n📝 Colunas extras (irão para ANOTAÇÕES): {preview['nao_mapeadas_origem']}")
    
    print("\n🔄 Convertendo...")
    resultado = converter_arquivo(
        args.arquivo, args.tipo, args.origem,
        id_inicial=args.id_inicial,
        estilo_caixa=estilo_caixa,
        df_cadastro_ref=df_cadastro_ref,
        df_projetos_ref=df_projetos_ref,
    )
    
    df_resultado = resultado['dados']
    alertas = resultado['alertas']
    
    # Mostrar alertas de padronização
    if alertas:
        print(f"\n{'='*50}")
        print("⚠️  ALERTAS DE PADRONIZAÇÃO (valores novos no DOit):")
        for alerta in alertas:
            print(f"   {alerta}")
        print(f"{'='*50}")
    
    # Salvar
    df_resultado.to_excel(saida, index=False)
    print(f"\n✅ Arquivo convertido salvo: {saida}")
    print(f"   {len(df_resultado)} registros | {len(df_resultado.columns)} colunas")


if __name__ == '__main__':
    main()
