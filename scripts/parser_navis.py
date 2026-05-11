"""
Parser para relatórios do sistema Navis.

O Navis exporta relatórios com metadados no topo (título, filtros, data de impressão)
e os dados começam várias linhas abaixo. Além disso, o financeiro de previsão tem
formato multi-linha onde a classificação financeira pode continuar na linha seguinte.

Os cadastros (clientes, fornecedores, contatos) são exportados em formato "ficha",
onde cada registro ocupa múltiplas linhas com labels fixos em posições específicas.

Tipos de relatório suportados:
- Movimentos de Conta Corrente (financeiro_cc)
- Contas a Pagar e Receber - Em Aberto (financeiro_previsao)
- Consulta Projetos - Cadastro (projetos)
- Clientes (clientes)
- Fornecedores (fornecedores)
- Contatos (contatos)
- Aplicação de Horas (horas)
"""

import pandas as pd
import numpy as np
import re


def _detectar_cabecalho_navis(df_raw: pd.DataFrame, palavras_chave: list) -> int:
    """
    Detecta a linha de cabeçalho em um relatório Navis.
    Procura a linha com maior score de palavras-chave.
    """
    melhor_linha = 0
    melhor_score = 0
    
    for i in range(min(20, len(df_raw))):
        row_vals = [str(v).lower().strip() for v in df_raw.iloc[i] if pd.notna(v) and str(v).strip()]
        score = sum(1 for val in row_vals for kw in palavras_chave if kw in val)
        score += len(row_vals) * 0.3
        if score > melhor_score:
            melhor_score = score
            melhor_linha = i
    
    return melhor_linha


def _detectar_tipo_relatorio(df_raw: pd.DataFrame) -> str:
    """
    Detecta automaticamente o tipo de relatório Navis com base no conteúdo.
    Retorna: 'financeiro_cc', 'financeiro_previsao', 'projetos', 'clientes',
             'fornecedores', 'contatos' ou 'horas'
    """
    # Verificar as primeiras 15 linhas por palavras-chave
    texto_topo = ''
    for i in range(min(15, len(df_raw))):
        row_vals = [str(v).lower().strip() for v in df_raw.iloc[i] if pd.notna(v)]
        texto_topo += ' '.join(row_vals) + ' '
    
    if 'conta corrente' in texto_topo or 'movimentos de conta' in texto_topo:
        return 'financeiro_cc'
    elif 'contas a pagar' in texto_topo or 'em aberto' in texto_topo:
        return 'financeiro_previsao'
    elif 'consulta projetos' in texto_topo or ('cadastro' in texto_topo and 'projeto' in texto_topo):
        return 'projetos'
    elif 'aplicação de horas' in texto_topo or 'aplicacao de horas' in texto_topo:
        return 'horas'
    elif 'listagem de contatos' in texto_topo:
        return 'contatos'
    elif 'clientes' in texto_topo and 'fornecedor' not in texto_topo:
        return 'clientes'
    elif 'fornecedores' in texto_topo or 'fornecedor' in texto_topo:
        return 'fornecedores'
    
    # Fallback: verificar colunas do cabeçalho
    for i in range(min(15, len(df_raw))):
        row_vals = [str(v).lower().strip() for v in df_raw.iloc[i] if pd.notna(v)]
        row_text = ' '.join(row_vals)
        if 'lcdo' in row_text and 'movimento' in row_text:
            return 'financeiro_cc'
        elif 'emissão' in row_text and 'vencto' in row_text:
            return 'financeiro_previsao'
        elif 'nome do projeto' in row_text:
            return 'projetos'
        elif 'cliente :' in row_text:
            return 'clientes'
        elif 'fornecedor :' in row_text:
            return 'fornecedores'
    
    return 'desconhecido'


def parse_navis_financeiro_cc(arquivo: str, sheet_name: str = None) -> pd.DataFrame:
    """
    Parseia relatório 'Movimentos de Conta Corrente' do Navis.
    
    Formato:
    - Linhas 0-8: metadados (título, data impressão, página)
    - Linha 9: cabeçalho (Lcdo, C/C, R/D, Movimento, Conciliado, Conciliação, 
               Bom Para, Nro. Doc., Pessoa, Descr./Nro. Cheque, Baixa/Lancto., 
               Valor, Juros, Descontos, Classificação Financeira, Observações)
    - Linhas 10+: dados
    """
    if isinstance(arquivo, str):
        xls = pd.ExcelFile(arquivo)
        if sheet_name is None:
            sheet_name = xls.sheet_names[0]
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    else:
        df_raw = pd.read_excel(arquivo, sheet_name=sheet_name or 0, header=None)
    
    # Detectar cabeçalho
    palavras_chave = ['lcdo', 'c/c', 'r/d', 'movimento', 'conciliado', 'conciliação',
                      'nro. doc', 'pessoa', 'descr', 'baixa', 'valor', 'juros',
                      'descontos', 'classificação']
    header_row = _detectar_cabecalho_navis(df_raw, palavras_chave)
    
    # Extrair cabeçalho - mapear posições das colunas
    header_vals = {}
    for j, v in enumerate(df_raw.iloc[header_row]):
        if pd.notna(v) and str(v).strip():
            header_vals[j] = str(v).strip()
    
    # Mapear colunas por posição conhecida do Navis CC
    # Baseado na análise: col 0=Lcdo, 5=C/C, 2=R/D, 4=Movimento, 6=Conciliado,
    # 7=Conciliação, 8=Bom Para, 9=Nro.Doc., 10=Pessoa, 11=Pessoa(tipo),
    # 12=Descr./Nro.Cheque, 13=Baixa/Lancto., 14=Valor, 15=Juros, 16=Descontos,
    # 20=Classificação Financeira, 21=Observações
    
    registros = []
    for i in range(header_row + 1, len(df_raw)):
        row = df_raw.iloc[i]
        
        # Pular linhas de rodapé (empresa, endereço, telefone)
        val_col0 = row.iloc[0] if pd.notna(row.iloc[0]) else ''
        val_col0_str = str(val_col0).strip()
        
        # Detectar se é uma linha de dados (tem valor na coluna Lcdo ou Movimento)
        lcdo = row.iloc[0] if pd.notna(row.iloc[0]) else None
        rd = row.iloc[2] if len(row) > 2 and pd.notna(row.iloc[2]) else None
        movimento = row.iloc[4] if len(row) > 4 and pd.notna(row.iloc[4]) else None
        valor = row.iloc[15] if len(row) > 15 and pd.notna(row.iloc[15]) else None
        
        # Linha de dados: tem Lcdo (S/N) e valor numérico
        if lcdo in ('S', 'N') and valor is not None:
            registro = {
                'Lcdo': lcdo,
                'R/D': rd,
                'Movimento': movimento,
                'Conciliado': row.iloc[7] if len(row) > 7 and pd.notna(row.iloc[7]) else None,
                'Conciliação': row.iloc[8] if len(row) > 8 and pd.notna(row.iloc[8]) else None,
                'Bom Para': row.iloc[9] if len(row) > 9 and pd.notna(row.iloc[9]) else None,
                'Nro. Doc.': row.iloc[10] if len(row) > 10 and pd.notna(row.iloc[10]) else None,
                'Pessoa': row.iloc[11] if len(row) > 11 and pd.notna(row.iloc[11]) else None,
                'Tipo Pessoa': row.iloc[12] if len(row) > 12 and pd.notna(row.iloc[12]) else None,
                'Descr. / Nro. Cheque': row.iloc[13] if len(row) > 13 and pd.notna(row.iloc[13]) else None,
                'Baixa/Lancto.': row.iloc[14] if len(row) > 14 and pd.notna(row.iloc[14]) else None,
                'Valor': valor,
                'Juros': row.iloc[16] if len(row) > 16 and pd.notna(row.iloc[16]) else 0,
                'Descontos': row.iloc[17] if len(row) > 17 and pd.notna(row.iloc[17]) else 0,
                'Classificação Financeira': row.iloc[20] if len(row) > 20 and pd.notna(row.iloc[20]) else None,
                'Observações': row.iloc[21] if len(row) > 21 and pd.notna(row.iloc[21]) else None,
            }
            registros.append(registro)
    
    df = pd.DataFrame(registros)
    
    if df.empty:
        return df
    
    # Converter Movimento para data
    df['Movimento'] = pd.to_datetime(df['Movimento'], errors='coerce')
    df['Conciliação'] = pd.to_datetime(df['Conciliação'], errors='coerce')
    
    # Converter R/D para Receita/Despesa
    df['Tipo'] = df['R/D'].apply(
        lambda x: 'Receita' if str(x).strip().upper() == 'R' else ('Despesa' if str(x).strip().upper() == 'D' else x)
    )
    
    # Conciliado: S → Sim, N → Não
    df['Conciliado_Texto'] = df['Lcdo'].apply(
        lambda x: 'Sim' if str(x).strip().upper() == 'S' else 'Não'
    )
    
    # Extrair categoria (código + descrição da classificação financeira)
    def _extrair_categoria(val):
        if pd.isna(val) or str(val).strip() in ('', '-'):
            return ''
        return str(val).strip()
    
    df['Categoria'] = df['Classificação Financeira'].apply(_extrair_categoria)
    
    return df


def parse_navis_financeiro_previsao(arquivo: str, sheet_name: str = None) -> pd.DataFrame:
    """
    Parseia relatório 'Contas a Pagar e Receber - Em Aberto' do Navis.
    
    Formato multi-linha:
    - Linhas 0-9: metadados
    - Linha 10: cabeçalho (Emissão, Vencto., Favorecido/Sacado, Classificação Financeira,
                Nº Doc., V. Bruto, Valor Liq., R/D, Parc., Observações)
    - Linhas 11+: dados, onde a classificação pode continuar na linha seguinte
      (linha de continuação tem dados apenas na coluna 8 e/ou 26)
    """
    if isinstance(arquivo, str):
        xls = pd.ExcelFile(arquivo)
        if sheet_name is None:
            sheet_name = xls.sheet_names[0]
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    else:
        df_raw = pd.read_excel(arquivo, sheet_name=sheet_name or 0, header=None)
    
    # Detectar cabeçalho
    palavras_chave = ['emissão', 'vencto', 'favorecido', 'sacado', 'classificação',
                      'nº doc', 'v. bruto', 'valor liq', 'r/d', 'parc', 'observações']
    header_row = _detectar_cabecalho_navis(df_raw, palavras_chave)
    
    # Parsear dados multi-linha
    registros = []
    registro_atual = None
    
    for i in range(header_row + 1, len(df_raw)):
        row = df_raw.iloc[i]
        
        # Verificar se é uma linha de dados principal (tem data na coluna 2 ou valor na coluna 16)
        emissao = row.iloc[2] if len(row) > 2 and pd.notna(row.iloc[2]) else None
        valor_bruto = row.iloc[16] if len(row) > 16 and pd.notna(row.iloc[16]) else None
        
        # Verificar se é linha de continuação (só tem classificação na col 8 e/ou obs na col 26)
        classificacao_cont = row.iloc[8] if len(row) > 8 and pd.notna(row.iloc[8]) else None
        obs_cont = row.iloc[26] if len(row) > 26 and pd.notna(row.iloc[26]) else None
        
        # Contar quantas colunas têm dados
        n_preenchidas = sum(1 for v in row if pd.notna(v) and str(v).strip())
        
        if emissao is not None and valor_bruto is not None:
            # Nova linha de dados principal
            if registro_atual is not None:
                registros.append(registro_atual)
            
            registro_atual = {
                'Emissão': emissao,
                'Vencimento': row.iloc[6] if len(row) > 6 and pd.notna(row.iloc[6]) else None,
                'Favorecido / Sacado': row.iloc[7] if len(row) > 7 and pd.notna(row.iloc[7]) else None,
                'Classificação Financeira': str(row.iloc[8]).strip() if len(row) > 8 and pd.notna(row.iloc[8]) else '',
                'Nº Doc.': row.iloc[12] if len(row) > 12 and pd.notna(row.iloc[12]) else None,
                'V. Bruto': valor_bruto,
                'Valor Liq.': row.iloc[20] if len(row) > 20 and pd.notna(row.iloc[20]) else valor_bruto,
                'R/D': row.iloc[22] if len(row) > 22 and pd.notna(row.iloc[22]) else None,
                'Parc.': row.iloc[24] if len(row) > 24 and pd.notna(row.iloc[24]) else None,
                'Observações': str(row.iloc[26]).strip() if len(row) > 26 and pd.notna(row.iloc[26]) else '',
            }
        
        elif registro_atual is not None and n_preenchidas <= 3 and classificacao_cont is not None:
            # Linha de continuação - concatenar classificação e observações
            registro_atual['Classificação Financeira'] += ' ' + str(classificacao_cont).strip()
            if obs_cont:
                registro_atual['Observações'] += ' ' + str(obs_cont).strip()
    
    # Não esquecer o último registro
    if registro_atual is not None:
        registros.append(registro_atual)
    
    df = pd.DataFrame(registros)
    
    if df.empty:
        return df
    
    # Converter datas
    df['Emissão'] = pd.to_datetime(df['Emissão'], dayfirst=True, errors='coerce')
    df['Vencimento'] = pd.to_datetime(df['Vencimento'], dayfirst=True, errors='coerce')
    
    # Converter R/D para Receita/Despesa
    df['Tipo'] = df['R/D'].apply(
        lambda x: 'Receita' if pd.notna(x) and str(x).strip().upper() == 'R' 
        else ('Despesa' if pd.notna(x) and str(x).strip().upper() == 'D' else x)
    )
    
    # Limpar classificação financeira (remover espaços extras)
    df['Classificação Financeira'] = df['Classificação Financeira'].str.strip()
    df['Observações'] = df['Observações'].str.strip()
    
    return df


def parse_navis_projetos(arquivo: str, sheet_name: str = None) -> pd.DataFrame:
    """
    Parseia relatório 'Consulta Projetos - Cadastro' do Navis.
    
    Formato:
    - Linhas 0-8: metadados
    - Linha 9: cabeçalho (Nome do Projeto, Nome do Cliente, Nome Fantasia,
               Tipo de Projeto, Status, Endereço, Cidade, UF, Dt.Início,
               Dt.Término, CM, Área, Observações, etc.)
    - Linhas 10+: dados (uma linha por projeto)
    """
    if isinstance(arquivo, str):
        xls = pd.ExcelFile(arquivo)
        if sheet_name is None:
            sheet_name = xls.sheet_names[0]
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    else:
        df_raw = pd.read_excel(arquivo, sheet_name=sheet_name or 0, header=None)
    
    # Detectar cabeçalho
    palavras_chave = ['nome do projeto', 'nome do cliente', 'nome fantasia',
                      'tipo de projeto', 'status', 'endereço', 'cidade', 'uf',
                      'dt.início', 'dt.término', 'área', 'observações']
    header_row = _detectar_cabecalho_navis(df_raw, palavras_chave)
    
    # Mapear posições das colunas do cabeçalho
    # Baseado na análise: col 0=Nome do Projeto, 5=Nome do Cliente, 6=Nome Fantasia,
    # 7=Tipo de Projeto, 10=Status, 11=Endereço, 14=Cidade, 15=UF,
    # 16=Dt.Início, 17=Dt.Término, 18=CM, 19=Área, 31=Observações
    
    registros = []
    for i in range(header_row + 1, len(df_raw)):
        row = df_raw.iloc[i]
        
        # Verificar se é linha de dados (tem nome do projeto na col 0)
        nome_projeto = row.iloc[0] if pd.notna(row.iloc[0]) else None
        if nome_projeto is None or str(nome_projeto).strip() == '':
            continue
        
        # Pular linhas de rodapé
        nome_str = str(nome_projeto).strip().lower()
        if any(x in nome_str for x in ['impresso:', 'página:', 'filtros:', 'total']):
            continue
        
        registro = {
            'Nome do Projeto': str(nome_projeto).strip(),
            'Nome do Cliente': str(row.iloc[5]).strip() if len(row) > 5 and pd.notna(row.iloc[5]) else '',
            'Nome Fantasia': str(row.iloc[6]).strip() if len(row) > 6 and pd.notna(row.iloc[6]) else '',
            'Tipo de Projeto': str(row.iloc[7]).strip() if len(row) > 7 and pd.notna(row.iloc[7]) else '',
            'Status': str(row.iloc[10]).strip() if len(row) > 10 and pd.notna(row.iloc[10]) else '',
            'Endereço': str(row.iloc[11]).strip() if len(row) > 11 and pd.notna(row.iloc[11]) else '',
            'Cidade': str(row.iloc[14]).strip() if len(row) > 14 and pd.notna(row.iloc[14]) else '',
            'UF': str(row.iloc[15]).strip() if len(row) > 15 and pd.notna(row.iloc[15]) else '',
            'Dt.Início': row.iloc[16] if len(row) > 16 and pd.notna(row.iloc[16]) else None,
            'Dt.Término': row.iloc[17] if len(row) > 17 and pd.notna(row.iloc[17]) else None,
            'CM': str(row.iloc[18]).strip() if len(row) > 18 and pd.notna(row.iloc[18]) else '',
            'Área': row.iloc[19] if len(row) > 19 and pd.notna(row.iloc[19]) else None,
            'Rotulo Base': str(row.iloc[21]).strip() if len(row) > 21 and pd.notna(row.iloc[21]) else '',
            'Observações': str(row.iloc[31]).strip() if len(row) > 31 and pd.notna(row.iloc[31]) else '',
        }
        registros.append(registro)
    
    df = pd.DataFrame(registros)
    
    if df.empty:
        return df
    
    # Converter datas
    df['Dt.Início'] = pd.to_datetime(df['Dt.Início'], errors='coerce')
    df['Dt.Término'] = pd.to_datetime(df['Dt.Término'], errors='coerce')
    
    return df


def parse_navis_clientes(arquivo: str, sheet_name: str = None) -> pd.DataFrame:
    """
    Parseia relatório de Clientes do Navis (formato ficha).
    
    Cada cliente ocupa um bloco de ~11 linhas:
    - "Cliente :": col 5=ID, col 6=Nome
    - "Endereço :": col 5=Endereço
    - "Bairro :": col 5=Bairro
    - "Cidade :": col 5=Cidade, col 9=UF
    - "CEP :": col 5=CEP, col 9=Email
    - "Fones :": col 5=Telefone, CNPJ na linha seguinte col 10
    """
    if isinstance(arquivo, str):
        xls = pd.ExcelFile(arquivo)
        if sheet_name is None:
            sheet_name = xls.sheet_names[0]
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    else:
        df_raw = pd.read_excel(arquivo, sheet_name=sheet_name or 0, header=None)
    
    registros = []
    i = 0
    
    while i < len(df_raw):
        row = df_raw.iloc[i]
        
        # Detectar início de registro: "Cliente :" na coluna 2
        is_cliente = False
        for j in range(min(5, len(row))):
            if pd.notna(row.iloc[j]) and 'cliente' in str(row.iloc[j]).lower() and ':' in str(row.iloc[j]):
                is_cliente = True
                break
        
        if is_cliente:
            registro = {
                'ID': None, 'Nome': '', 'Endereço': '', 'Bairro': '',
                'Cidade': '', 'Estado': '', 'CEP': '', 'Email': '',
                'Telefone': '', 'CNPJ/CPF': '', 'IE/RG': '', 'Contato': '',
            }
            
            # ID e Nome na mesma linha
            registro['ID'] = row.iloc[5] if len(row) > 5 and pd.notna(row.iloc[5]) else None
            registro['Nome'] = str(row.iloc[6]).strip() if len(row) > 6 and pd.notna(row.iloc[6]) else ''
            
            # Percorrer próximas linhas do bloco
            for offset in range(1, 15):
                if i + offset >= len(df_raw):
                    break
                next_row = df_raw.iloc[i + offset]
                
                # Verificar se é início de novo registro
                is_next_cliente = False
                for j in range(min(5, len(next_row))):
                    if pd.notna(next_row.iloc[j]) and 'cliente' in str(next_row.iloc[j]).lower() and ':' in str(next_row.iloc[j]):
                        is_next_cliente = True
                        break
                if is_next_cliente:
                    break
                
                # Extrair dados por label
                row_text = ' '.join(str(v).lower().strip() for v in next_row if pd.notna(v))
                
                if 'endereço' in row_text and ':' in row_text:
                    registro['Endereço'] = str(next_row.iloc[5]).strip() if len(next_row) > 5 and pd.notna(next_row.iloc[5]) else ''
                
                elif 'bairro' in row_text and ':' in row_text:
                    registro['Bairro'] = str(next_row.iloc[5]).strip() if len(next_row) > 5 and pd.notna(next_row.iloc[5]) else ''
                    # Contato pode estar na col 8
                    if len(next_row) > 9 and pd.notna(next_row.iloc[9]):
                        registro['Contato'] = str(next_row.iloc[9]).strip()
                
                elif 'cidade' in row_text and ':' in row_text:
                    registro['Cidade'] = str(next_row.iloc[5]).strip() if len(next_row) > 5 and pd.notna(next_row.iloc[5]) else ''
                    # Estado na col 9
                    if len(next_row) > 9 and pd.notna(next_row.iloc[9]):
                        registro['Estado'] = str(next_row.iloc[9]).strip()
                
                elif 'cep' in row_text and ':' in row_text:
                    registro['CEP'] = str(next_row.iloc[5]).strip() if len(next_row) > 5 and pd.notna(next_row.iloc[5]) else ''
                    # Email na col 9
                    if len(next_row) > 9 and pd.notna(next_row.iloc[9]):
                        registro['Email'] = str(next_row.iloc[9]).strip()
                
                elif 'fones' in row_text and ':' in row_text:
                    # Telefone na col 5
                    if len(next_row) > 5 and pd.notna(next_row.iloc[5]) and 'fones' not in str(next_row.iloc[5]).lower():
                        registro['Telefone'] = str(next_row.iloc[5]).strip()
                    # CNPJ pode estar na col 10 da mesma linha ou na próxima
                    if len(next_row) > 10 and pd.notna(next_row.iloc[10]):
                        val = str(next_row.iloc[10]).strip()
                        if val and val not in ('CNPJ / CPF :', 'CNPJ / CPF :'):
                            registro['CNPJ/CPF'] = val
                
                elif not registro['CNPJ/CPF']:
                    # CNPJ pode estar na linha seguinte ao Fones
                    if len(next_row) > 10 and pd.notna(next_row.iloc[10]):
                        val = str(next_row.iloc[10]).strip()
                        if re.match(r'[\d./-]+', val) and len(val) > 10:
                            registro['CNPJ/CPF'] = val
                
                if 'ie / rg' in row_text and ':' in row_text:
                    if len(next_row) > 10 and pd.notna(next_row.iloc[10]):
                        val = str(next_row.iloc[10]).strip()
                        if val and 'ie' not in val.lower():
                            registro['IE/RG'] = val
            
            registros.append(registro)
        
        i += 1
    
    return pd.DataFrame(registros)


def parse_navis_fornecedores(arquivo: str, sheet_name: str = None) -> pd.DataFrame:
    """
    Parseia relatório de Fornecedores do Navis (formato ficha).
    
    Cada fornecedor ocupa um bloco de ~19 linhas:
    - "Fornecedor :": col 6=ID, col 7=Nome
    - Linha seguinte: col 6=Nome Fantasia
    - "Endereço :": col 6=Endereço
    - "Bairro :": col 6=Bairro
    - "Cidade :": col 6=Cidade, col 12=UF
    - "CEP :": col 6=CEP, col 12=Email
    - "Fones :": col 12=CNPJ
    - Linha telefone: col 6=Telefone
    - "Atividade:": col 9=Atividade
    - "Contatos:": seguida de nomes
    """
    if isinstance(arquivo, str):
        xls = pd.ExcelFile(arquivo)
        if sheet_name is None:
            sheet_name = xls.sheet_names[0]
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    else:
        df_raw = pd.read_excel(arquivo, sheet_name=sheet_name or 0, header=None)
    
    registros = []
    i = 0
    
    while i < len(df_raw):
        row = df_raw.iloc[i]
        
        # Detectar início de registro: "Fornecedor :" 
        is_fornecedor = False
        for j in range(min(5, len(row))):
            if pd.notna(row.iloc[j]) and 'fornecedor' in str(row.iloc[j]).lower() and ':' in str(row.iloc[j]):
                is_fornecedor = True
                break
        
        if is_fornecedor:
            registro = {
                'ID': None, 'Nome': '', 'Nome Fantasia': '', 'Endereço': '',
                'Bairro': '', 'Cidade': '', 'Estado': '', 'CEP': '',
                'Email': '', 'Telefone': '', 'CNPJ/CPF': '', 'IE/RG': '',
                'Atividade': '', 'Contatos': '',
            }
            
            # ID e Nome
            registro['ID'] = row.iloc[6] if len(row) > 6 and pd.notna(row.iloc[6]) else None
            registro['Nome'] = str(row.iloc[7]).strip() if len(row) > 7 and pd.notna(row.iloc[7]) else ''
            
            # Nome Fantasia na linha seguinte, col 6
            if i + 1 < len(df_raw):
                next_r = df_raw.iloc[i + 1]
                if len(next_r) > 6 and pd.notna(next_r.iloc[6]):
                    val = str(next_r.iloc[6]).strip()
                    # Verificar que não é "Endereço :" ou outro label
                    if val and 'endereço' not in val.lower():
                        registro['Nome Fantasia'] = val
            
            # Percorrer próximas linhas do bloco
            encontrou_fones = False
            for offset in range(1, 25):
                if i + offset >= len(df_raw):
                    break
                next_row = df_raw.iloc[i + offset]
                
                # Verificar se é início de novo registro
                is_next = False
                for j in range(min(5, len(next_row))):
                    if pd.notna(next_row.iloc[j]) and 'fornecedor' in str(next_row.iloc[j]).lower() and ':' in str(next_row.iloc[j]):
                        is_next = True
                        break
                if is_next:
                    break
                
                row_text = ' '.join(str(v).lower().strip() for v in next_row if pd.notna(v))
                
                if 'endereço' in row_text and ':' in row_text:
                    registro['Endereço'] = str(next_row.iloc[6]).strip() if len(next_row) > 6 and pd.notna(next_row.iloc[6]) else ''
                
                elif 'bairro' in row_text and ':' in row_text and not encontrou_fones:
                    registro['Bairro'] = str(next_row.iloc[6]).strip() if len(next_row) > 6 and pd.notna(next_row.iloc[6]) else ''
                
                elif 'cidade' in row_text and ':' in row_text and not encontrou_fones:
                    registro['Cidade'] = str(next_row.iloc[6]).strip() if len(next_row) > 6 and pd.notna(next_row.iloc[6]) else ''
                    if len(next_row) > 12 and pd.notna(next_row.iloc[12]):
                        registro['Estado'] = str(next_row.iloc[12]).strip()
                
                elif 'cep' in row_text and ':' in row_text and not encontrou_fones:
                    registro['CEP'] = str(next_row.iloc[6]).strip() if len(next_row) > 6 and pd.notna(next_row.iloc[6]) else ''
                    if len(next_row) > 12 and pd.notna(next_row.iloc[12]):
                        registro['Email'] = str(next_row.iloc[12]).strip()
                
                elif 'fones' in row_text and ':' in row_text:
                    encontrou_fones = True
                    # CNPJ na col 12
                    if len(next_row) > 12 and pd.notna(next_row.iloc[12]):
                        val = str(next_row.iloc[12]).strip()
                        if val and 'cnpj' not in val.lower():
                            registro['CNPJ/CPF'] = val
                
                elif 'ie / rg' in row_text and ':' in row_text:
                    if len(next_row) > 12 and pd.notna(next_row.iloc[12]):
                        val = str(next_row.iloc[12]).strip()
                        if val and 'ie' not in val.lower():
                            registro['IE/RG'] = val
                
                elif 'atividade' in row_text and ':' in row_text:
                    if len(next_row) > 9 and pd.notna(next_row.iloc[9]):
                        registro['Atividade'] = str(next_row.iloc[9]).strip().rstrip(' -')
                
                elif 'contatos' in row_text and ':' in row_text:
                    # Próxima linha pode ter nome do contato
                    if i + offset + 1 < len(df_raw):
                        contato_row = df_raw.iloc[i + offset + 1]
                        if len(contato_row) > 1 and pd.notna(contato_row.iloc[1]):
                            val = str(contato_row.iloc[1]).strip()
                            if val and 'fornecedor' not in val.lower():
                                registro['Contatos'] = val
                
                elif encontrou_fones and not registro['Telefone']:
                    # Telefone pode estar na col 6 após a linha de Fones
                    if len(next_row) > 6 and pd.notna(next_row.iloc[6]):
                        val = str(next_row.iloc[6]).strip()
                        if re.search(r'\(\d{2}\)', val) or re.match(r'[\d\s\-]+$', val):
                            registro['Telefone'] = val
            
            registros.append(registro)
        
        i += 1
    
    return pd.DataFrame(registros)


def parse_navis_contatos(arquivo: str, sheet_name: str = None) -> pd.DataFrame:
    """
    Parseia relatório de Contatos do Navis (formato ficha).
    
    Cada contato ocupa um bloco de ~16 linhas:
    - "Código:": col 7=ID, col 10=Nome
    - "Data Nasc:": col 8="CREA:", col 17="E-Mail:", col 18=Email
    - "Endereço :": col 6=Endereço
    - "Bairro :": col 6=Bairro
    - "Complemento :": col 27=Complemento
    - "Cidade :": col 6=Cidade, col 18=UF
    - "CEP :": col 6=CEP
    - "CNPJ / CPF :": col seguinte com valor
    - Telefone na linha após Fones
    """
    if isinstance(arquivo, str):
        xls = pd.ExcelFile(arquivo)
        if sheet_name is None:
            sheet_name = xls.sheet_names[0]
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    else:
        df_raw = pd.read_excel(arquivo, sheet_name=sheet_name or 0, header=None)
    
    registros = []
    i = 0
    
    while i < len(df_raw):
        row = df_raw.iloc[i]
        
        # Detectar início de registro: "Código:" na coluna 3
        is_contato = False
        for j in range(min(6, len(row))):
            if pd.notna(row.iloc[j]) and 'código' in str(row.iloc[j]).lower() and ':' in str(row.iloc[j]):
                is_contato = True
                break
        
        if is_contato:
            registro = {
                'ID': None, 'Nome': '', 'Email': '', 'Endereço': '',
                'Complemento': '', 'Bairro': '', 'Cidade': '', 'Estado': '',
                'CEP': '', 'Telefone': '', 'CNPJ/CPF': '', 'IE/RG': '',
                'Site': '', 'CREA': '',
            }
            
            # ID e Nome
            registro['ID'] = row.iloc[7] if len(row) > 7 and pd.notna(row.iloc[7]) else None
            registro['Nome'] = str(row.iloc[10]).strip() if len(row) > 10 and pd.notna(row.iloc[10]) else ''
            
            # Percorrer próximas linhas do bloco
            for offset in range(1, 20):
                if i + offset >= len(df_raw):
                    break
                next_row = df_raw.iloc[i + offset]
                
                # Verificar se é início de novo registro
                is_next = False
                for j in range(min(6, len(next_row))):
                    if pd.notna(next_row.iloc[j]) and 'código' in str(next_row.iloc[j]).lower() and ':' in str(next_row.iloc[j]):
                        is_next = True
                        break
                if is_next:
                    break
                
                row_text = ' '.join(str(v).lower().strip() for v in next_row if pd.notna(v))
                
                if 'data nasc' in row_text or 'crea' in row_text:
                    # Email na col 18
                    if len(next_row) > 18 and pd.notna(next_row.iloc[18]):
                        registro['Email'] = str(next_row.iloc[18]).strip()
                
                elif 'endereço' in row_text and ':' in row_text:
                    registro['Endereço'] = str(next_row.iloc[6]).strip() if len(next_row) > 6 and pd.notna(next_row.iloc[6]) else ''
                
                elif 'site' in row_text and ':' in row_text:
                    if len(next_row) > 16 and pd.notna(next_row.iloc[16]):
                        registro['Site'] = str(next_row.iloc[16]).strip()
                
                elif 'bairro' in row_text and ':' in row_text:
                    registro['Bairro'] = str(next_row.iloc[6]).strip() if len(next_row) > 6 and pd.notna(next_row.iloc[6]) else ''
                
                elif 'complemento' in row_text and ':' in row_text:
                    if len(next_row) > 27 and pd.notna(next_row.iloc[27]):
                        registro['Complemento'] = str(next_row.iloc[27]).strip()
                
                elif 'cidade' in row_text and ':' in row_text:
                    registro['Cidade'] = str(next_row.iloc[6]).strip() if len(next_row) > 6 and pd.notna(next_row.iloc[6]) else ''
                    if len(next_row) > 18 and pd.notna(next_row.iloc[18]):
                        registro['Estado'] = str(next_row.iloc[18]).strip()
                
                elif 'cep' in row_text and ':' in row_text:
                    registro['CEP'] = str(next_row.iloc[6]).strip() if len(next_row) > 6 and pd.notna(next_row.iloc[6]) else ''
                
                elif 'cnpj' in row_text and 'cpf' in row_text and ':' in row_text:
                    # CNPJ pode estar na mesma linha ou na próxima
                    for col_idx in range(12, min(20, len(next_row))):
                        if pd.notna(next_row.iloc[col_idx]):
                            val = str(next_row.iloc[col_idx]).strip()
                            if val and 'cnpj' not in val.lower() and 'cpf' not in val.lower():
                                registro['CNPJ/CPF'] = val
                                break
                
                elif 'ie / rg' in row_text and ':' in row_text:
                    for col_idx in range(12, min(20, len(next_row))):
                        if pd.notna(next_row.iloc[col_idx]):
                            val = str(next_row.iloc[col_idx]).strip()
                            if val and 'ie' not in val.lower() and 'rg' not in val.lower():
                                registro['IE/RG'] = val
                                break
                
                elif 'fones' in row_text and ':' in row_text:
                    pass  # Telefone na próxima iteração
                
                elif not registro['Telefone']:
                    # Procurar telefone (formato (XX) XXXX-XXXX ou números)
                    for col_idx in range(5, min(15, len(next_row))):
                        if pd.notna(next_row.iloc[col_idx]):
                            val = str(next_row.iloc[col_idx]).strip()
                            if re.search(r'\d{4,}', val) and 'banco' not in row_text and 'agência' not in row_text:
                                registro['Telefone'] = val
                                break
            
            registros.append(registro)
        
        i += 1
    
    return pd.DataFrame(registros)


def parse_navis_horas(arquivo: str, sheet_name: str = None) -> pd.DataFrame:
    """
    Parseia relatório de Aplicação de Horas do Navis.
    
    Formato:
    - Cabeçalho com "APLICAÇÃO DE HORAS", usuário e período
    - Blocos por semana com "Semana: X / Y"
    - Dentro de cada semana: linhas com Data (col 8), Horas (col 9), Projeto (col 12)
    - Subtotal de horas por semana
    
    O nome do colaborador é extraído do campo "Usuário:" no cabeçalho.
    """
    if isinstance(arquivo, str):
        xls = pd.ExcelFile(arquivo)
        if sheet_name is None:
            sheet_name = xls.sheet_names[0]
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    else:
        df_raw = pd.read_excel(arquivo, sheet_name=sheet_name or 0, header=None)
    
    # Extrair nome do colaborador do cabeçalho
    colaborador = ''
    for i in range(min(10, len(df_raw))):
        row = df_raw.iloc[i]
        for j in range(len(row)):
            if pd.notna(row.iloc[j]) and 'usuário' in str(row.iloc[j]).lower():
                val = str(row.iloc[j]).strip()
                # Formato: "Usuário:NOME"
                if ':' in val:
                    colaborador = val.split(':', 1)[1].strip()
                break
    
    # Extrair registros de horas
    registros = []
    
    for i in range(len(df_raw)):
        row = df_raw.iloc[i]
        
        # Detectar linha de dados: tem data na col 8 e horas na col 9
        data_val = row.iloc[8] if len(row) > 8 and pd.notna(row.iloc[8]) else None
        horas_val = row.iloc[9] if len(row) > 9 and pd.notna(row.iloc[9]) else None
        projeto_val = row.iloc[12] if len(row) > 12 and pd.notna(row.iloc[12]) else None
        
        # Pular linhas de cabeçalho "Data" / "Horas" / "Projeto"
        if data_val and str(data_val).strip().lower() == 'data':
            continue
        
        # Pular linhas de subtotal
        row_text = ' '.join(str(v).lower() for v in row if pd.notna(v))
        if 'subtotal' in row_text or 'total' in row_text:
            continue
        
        # Verificar se é uma linha de dados válida
        if data_val is not None and horas_val is not None and projeto_val is not None:
            # Converter horas (formato HH:MM:SS ou número)
            horas_decimal = None
            horas_str = str(horas_val).strip()
            if ':' in horas_str:
                partes = horas_str.split(':')
                try:
                    h = int(partes[0])
                    m = int(partes[1]) if len(partes) > 1 else 0
                    horas_decimal = h + m / 60.0
                except ValueError:
                    horas_decimal = None
            else:
                try:
                    horas_decimal = float(horas_str)
                except ValueError:
                    horas_decimal = None
            
            registro = {
                'Data': data_val,
                'Horas': horas_decimal,
                'Horas Original': horas_str,
                'Projeto': str(projeto_val).strip(),
                'Colaborador': colaborador,
            }
            registros.append(registro)
    
    df = pd.DataFrame(registros)
    
    if df.empty:
        return df
    
    # Converter data
    df['Data'] = pd.to_datetime(df['Data'], dayfirst=True, errors='coerce')
    
    return df


def parse_navis(arquivo: str, sheet_name: str = None, tipo_relatorio: str = None) -> pd.DataFrame:
    """
    Parser principal do Navis. Detecta automaticamente o tipo de relatório
    ou usa o tipo informado.
    
    Args:
        arquivo: Caminho do arquivo .xlsx
        sheet_name: Nome da aba (None = primeira aba)
        tipo_relatorio: 'financeiro_cc', 'financeiro_previsao', 'projetos',
                       'clientes', 'fornecedores', 'contatos' ou 'horas'
                       Se None, detecta automaticamente.
    
    Returns:
        DataFrame com os dados parseados
    """
    if isinstance(arquivo, str):
        xls = pd.ExcelFile(arquivo)
        if sheet_name is None:
            sheet_name = xls.sheet_names[0]
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=15)
    else:
        df_raw = pd.read_excel(arquivo, sheet_name=sheet_name or 0, header=None, nrows=15)
    
    # Detectar tipo se não informado
    if tipo_relatorio is None:
        tipo_relatorio = _detectar_tipo_relatorio(df_raw)
    
    if tipo_relatorio == 'financeiro_cc':
        return parse_navis_financeiro_cc(arquivo, sheet_name)
    elif tipo_relatorio == 'financeiro_previsao':
        return parse_navis_financeiro_previsao(arquivo, sheet_name)
    elif tipo_relatorio == 'projetos':
        return parse_navis_projetos(arquivo, sheet_name)
    elif tipo_relatorio == 'clientes':
        return parse_navis_clientes(arquivo, sheet_name)
    elif tipo_relatorio == 'fornecedores':
        return parse_navis_fornecedores(arquivo, sheet_name)
    elif tipo_relatorio == 'contatos':
        return parse_navis_contatos(arquivo, sheet_name)
    elif tipo_relatorio == 'horas':
        return parse_navis_horas(arquivo, sheet_name)
    else:
        raise ValueError(
            f"Tipo de relatório Navis não reconhecido: '{tipo_relatorio}'. "
            f"Use: 'financeiro_cc', 'financeiro_previsao', 'projetos', "
            f"'clientes', 'fornecedores', 'contatos' ou 'horas'"
        )
