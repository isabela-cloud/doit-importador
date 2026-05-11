"""
Parser para Financeiro Horizontal (BPO / Fluxo de Caixa)
Converte planilhas com meses lado a lado em lista de lançamentos.

Formato esperado:
- Colunas repetidas por mês: DESCRIÇÃO | VENC. | PAGO | PREVISTO | REALIZADO | OBSERVAÇÃO
- Seções: RECEITAS (Novos Contratos, Fluxo Pagtos, Extras), DESPESAS (Admin, Equipe, etc.)
"""

import pandas as pd
import numpy as np
import re
from datetime import datetime


# Meses em português e inglês para detecção
MESES_PT = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
            'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro']
MESES_EN = ['jan', 'feb', 'mar', 'apr', 'may', 'jun',
            'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
MESES_NUM = {m: i+1 for i, m in enumerate(MESES_PT)}
MESES_NUM.update({m: i+1 for i, m in enumerate(MESES_EN)})

# Categorias de seção → mapeamento para plano de contas DOit
SECAO_PARA_CATEGORIA = {
    'novos contratos': ('Receitas', 'Receita de Projeto', 'Etapa/Parcela'),
    'fluxo pag': ('Receitas', 'Receita de Projeto', 'Etapa/Parcela'),
    'fluxo de extras': ('Receitas', 'Reembolso de Despesas', ''),
    'despesas administrativas': ('Despesas Fixas', '', ''),
    'despesas equipe': ('Custos', 'Mão de Obra', ''),
    'marketing': ('Despesas Variáveis', 'Marketing', ''),
    'despesas cau': ('Despesas Fixas', 'Associações', ''),
    'tecnologia': ('Despesas Fixas', 'Software', ''),
    'montagem': ('Despesas Variáveis', 'Manutenção', ''),
    'despesas variav': ('Despesas Variáveis', '', ''),
}

# Palavras que indicam linhas de total/subtotal (ignorar)
PALAVRAS_TOTAL = ['total', 'subtotal', 'abatimento', 'reserva', 'lucro',
                  'comissão', 'receita', 'despesas totais', 'saldo']


def _detectar_ano(df_raw):
    """Tenta detectar o ano do BPO."""
    for i in range(min(5, len(df_raw))):
        for val in df_raw.iloc[i]:
            if pd.notna(val):
                val_str = str(val).strip()
                match = re.search(r'20\d{2}', val_str)
                if match:
                    return int(match.group())
    return datetime.now().year


def _detectar_meses_colunas(df_raw):
    """
    Detecta quais colunas correspondem a cada mês.
    Retorna dict: {mes_num: {'inicio': col_idx, 'fim': col_idx}}
    """
    meses_encontrados = {}
    
    # Procurar nas primeiras linhas por nomes de meses
    for i in range(min(5, len(df_raw))):
        for j, val in enumerate(df_raw.iloc[i]):
            if pd.notna(val):
                val_lower = str(val).strip().lower()
                for mes_nome, mes_num in MESES_NUM.items():
                    if mes_nome in val_lower and mes_num not in meses_encontrados:
                        meses_encontrados[mes_num] = {'inicio': j, 'label': val_lower}
                        break
    
    if not meses_encontrados:
        return {}
    
    # Ordenar e calcular fim de cada bloco
    meses_ordenados = sorted(meses_encontrados.keys())
    for i, mes in enumerate(meses_ordenados):
        if i + 1 < len(meses_ordenados):
            prox_mes = meses_ordenados[i + 1]
            meses_encontrados[mes]['fim'] = meses_encontrados[prox_mes]['inicio'] - 1
        else:
            meses_encontrados[mes]['fim'] = len(df_raw.columns) - 1
    
    return meses_encontrados


def _extrair_lancamentos_mes(df_raw, col_inicio, col_fim, mes_num, ano, secao_atual):
    """
    Extrai lançamentos de um bloco de mês.
    Cada bloco tem ~6 colunas: DESCRIÇÃO | VENC. | PAGO | PREVISTO | REALIZADO | OBSERVAÇÃO
    """
    lancamentos = []
    
    # O bloco de cada mês tem tipicamente 6 colunas
    # Coluna 0: Descrição/Nome
    # Coluna 1: Vencimento
    # Coluna 2: Pago (data)
    # Coluna 3: Previsto (valor)
    # Coluna 4: Realizado (valor)
    # Coluna 5: Observação
    
    num_colunas_bloco = col_fim - col_inicio + 1
    
    for idx, row in df_raw.iterrows():
        # Pegar valores do bloco deste mês
        valores_bloco = row.iloc[col_inicio:col_fim + 1].tolist()
        
        # Pegar o primeiro valor não-vazio como descrição
        descricao = None
        for v in valores_bloco[:2]:
            if pd.notna(v) and str(v).strip():
                descricao = str(v).strip()
                break
        
        if not descricao:
            continue
        
        # Ignorar linhas de total/subtotal/cabeçalho
        desc_lower = descricao.lower()
        if any(p in desc_lower for p in PALAVRAS_TOTAL):
            continue
        if any(m in desc_lower for m in MESES_PT):
            continue
        if desc_lower in ('receitas', 'despesas', '2026', '2025'):
            continue
        # Ignorar cabeçalhos de seção repetidos
        if desc_lower in ('venc.', 'pago', 'previsto', 'realizado', 'observação'):
            continue
        
        # Extrair valores
        vencimento = ''
        data_pago = ''
        valor_previsto = None
        valor_realizado = None
        observacao = ''
        
        if num_colunas_bloco >= 6:
            venc_raw = valores_bloco[1] if len(valores_bloco) > 1 else None
            pago_raw = valores_bloco[2] if len(valores_bloco) > 2 else None
            prev_raw = valores_bloco[3] if len(valores_bloco) > 3 else None
            real_raw = valores_bloco[4] if len(valores_bloco) > 4 else None
            obs_raw = valores_bloco[5] if len(valores_bloco) > 5 else None
        elif num_colunas_bloco >= 4:
            venc_raw = valores_bloco[1] if len(valores_bloco) > 1 else None
            pago_raw = None
            prev_raw = valores_bloco[2] if len(valores_bloco) > 2 else None
            real_raw = valores_bloco[3] if len(valores_bloco) > 3 else None
            obs_raw = None
        else:
            continue
        
        # Processar vencimento
        if pd.notna(venc_raw) and str(venc_raw).strip():
            vencimento = str(venc_raw).strip()
        
        # Processar data pago
        if pd.notna(pago_raw) and str(pago_raw).strip():
            data_pago = str(pago_raw).strip()
        
        # Processar valores
        def _parse_valor(v):
            if pd.isna(v) or str(v).strip() in ('', '-', 'R$ - 0', 'R$    - 0'):
                return None
            if isinstance(v, (int, float)):
                return float(v)
            val_str = str(v).strip().replace('R$', '').replace(' ', '')
            val_str = val_str.replace('.', '').replace(',', '.')
            try:
                return float(val_str)
            except ValueError:
                return None
        
        valor_previsto = _parse_valor(prev_raw)
        valor_realizado = _parse_valor(real_raw)
        
        # Usar realizado se disponível, senão previsto
        valor_final = valor_realizado if valor_realizado is not None else valor_previsto
        
        if valor_final is None or valor_final == 0:
            continue
        
        # Observação
        if pd.notna(obs_raw) and str(obs_raw).strip():
            observacao = str(obs_raw).strip()
        
        # Formatar data de vencimento
        data_formatada = ''
        if vencimento:
            try:
                # Tentar formatos comuns: "15-Jan", "27-Jan", "25/01/2026"
                for fmt in ['%d-%b', '%d/%m/%Y', '%d/%m/%y']:
                    try:
                        dt = pd.to_datetime(vencimento, format=fmt)
                        if dt.year < 2000:
                            dt = dt.replace(year=ano)
                        data_formatada = dt.strftime('%d/%m/%Y')
                        break
                    except (ValueError, TypeError):
                        continue
                if not data_formatada:
                    dt = pd.to_datetime(vencimento, dayfirst=True, errors='coerce')
                    if pd.notna(dt):
                        if dt.year < 2000:
                            dt = dt.replace(year=ano)
                        data_formatada = dt.strftime('%d/%m/%Y')
            except Exception:
                data_formatada = vencimento
        
        # Se não tem data de vencimento, usar mês/ano
        if not data_formatada:
            data_formatada = f"01/{mes_num:02d}/{ano}"
        
        # Determinar categoria
        cat1, cat2, cat3 = '', '', ''
        if secao_atual:
            for chave, cats in SECAO_PARA_CATEGORIA.items():
                if chave in secao_atual.lower():
                    cat1, cat2, cat3 = cats
                    break
        
        # Determinar tipo (receita/despesa)
        tipo = 'Receita' if valor_final > 0 else 'Despesa'
        if 'despesa' in secao_atual.lower() if secao_atual else False:
            tipo = 'Despesa'
            valor_final = -abs(valor_final)
        
        lancamentos.append({
            'DESCRIÇÃO': descricao,
            'VENCIMENTO': data_formatada,
            'DATA': data_pago if data_pago else '',
            'VALOR': valor_final,
            'TIPO': tipo,
            '1ª CATEGORIA': cat1,
            '2ª CATEGORIA': cat2,
            '3ª CATEGORIA': cat3,
            'CONCILIADO': 'SIM' if data_pago else 'NÃO',
            'OBSERVAÇÃO': observacao,
            '_MES': mes_num,
            '_SECAO': secao_atual or '',
        })
    
    return lancamentos


def parse_financeiro_horizontal(filepath, sheet_name=None, ano=None):
    """
    Parser principal para BPO / Financeiro Horizontal.
    
    Args:
        filepath: Caminho do arquivo Excel
        sheet_name: Nome da aba (None = primeira)
        ano: Ano de referência (None = detectar)
    
    Returns:
        DataFrame com lançamentos no formato vertical (uma linha por lançamento)
    """
    # Ler sem header
    if sheet_name:
        df_raw = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
    else:
        df_raw = pd.read_excel(filepath, header=None)
    
    # Detectar ano
    if not ano:
        ano = _detectar_ano(df_raw)
    
    # Detectar blocos de meses
    meses = _detectar_meses_colunas(df_raw)
    
    if not meses:
        raise ValueError("Não foi possível detectar os meses na planilha. Verifique se o formato é horizontal com meses lado a lado.")
    
    # Detectar seções e extrair lançamentos
    todos_lancamentos = []
    secao_atual = ''
    
    # Palavras que indicam início de seção
    secoes_conhecidas = [
        'novos contratos', 'fluxo pag', 'fluxo de extras',
        'despesas administrativas', 'despesas equipe', 'marketing e divulgação',
        'despesas cau', 'tecnologia da informação', 'montagem e instalação',
        'despesas variav',
    ]
    
    for idx, row in df_raw.iterrows():
        # Verificar se é uma linha de seção
        for col_val in row:
            if pd.notna(col_val):
                val_lower = str(col_val).strip().lower()
                for secao in secoes_conhecidas:
                    if secao in val_lower:
                        secao_atual = str(col_val).strip()
                        break
        
        # Para cada mês, tentar extrair lançamento desta linha
        for mes_num, info in meses.items():
            col_inicio = info['inicio']
            col_fim = info['fim']
            
            # Pegar valores do bloco
            if col_fim >= len(row):
                continue
            
            valores = row.iloc[col_inicio:col_fim + 1].tolist()
            
            # Encontrar descrição (primeiro texto não-vazio do bloco)
            descricao = None
            desc_col = None
            for i, v in enumerate(valores[:2]):
                if pd.notna(v) and str(v).strip():
                    val_str = str(v).strip()
                    # Ignorar se é número puro, data, ou valor monetário
                    if re.match(r'^[\d.,R$\s\-]+$', val_str):
                        continue
                    if val_str.lower() in ('venc.', 'pago', 'previsto', 'realizado', 'observação'):
                        continue
                    descricao = val_str
                    desc_col = i
                    break
            
            if not descricao:
                continue
            
            # Ignorar totais e cabeçalhos
            desc_lower = descricao.lower()
            if any(p in desc_lower for p in PALAVRAS_TOTAL):
                continue
            if any(m in desc_lower for m in MESES_PT + ['2026', '2025', '2024', 'receitas', 'despesas']):
                continue
            if desc_lower.startswith(('(-)','fluxo pag', 'fluxo de extra', 'novos contrato',
                                      'despesas admin', 'despesas equipe', 'marketing',
                                      'despesas cau', 'tecnologia', 'montagem', 'despesas variav')):
                continue
            
            # Extrair valores restantes do bloco
            num_cols = col_fim - col_inicio + 1
            
            # Tentar identificar: vencimento, pago, previsto, realizado, obs
            vencimento = ''
            data_pago = ''
            valor_previsto = None
            valor_realizado = None
            observacao = ''
            
            # Offset após a descrição
            offset = (desc_col or 0) + 1
            remaining = valores[offset:]
            
            # Procurar valores numéricos e datas nos campos restantes
            valores_numericos = []
            datas_encontradas = []
            textos = []
            
            for v in remaining:
                if pd.isna(v) or str(v).strip() == '':
                    continue
                v_str = str(v).strip()
                
                # É valor monetário?
                v_clean = v_str.replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.')
                try:
                    num = float(v_clean)
                    valores_numericos.append(num)
                    continue
                except ValueError:
                    pass
                
                # É número direto?
                if isinstance(v, (int, float)):
                    valores_numericos.append(float(v))
                    continue
                
                # É data?
                try:
                    dt = pd.to_datetime(v, dayfirst=True, errors='coerce')
                    if pd.notna(dt):
                        datas_encontradas.append(dt)
                        continue
                except Exception:
                    pass
                
                # É texto (observação)
                textos.append(v_str)
            
            # Atribuir valores encontrados
            if datas_encontradas:
                vencimento = datas_encontradas[0].strftime('%d/%m/%Y')
                if len(datas_encontradas) > 1:
                    data_pago = datas_encontradas[1].strftime('%d/%m/%Y')
            
            if valores_numericos:
                if len(valores_numericos) >= 2:
                    valor_previsto = valores_numericos[0]
                    valor_realizado = valores_numericos[1]
                else:
                    valor_realizado = valores_numericos[0]
            
            if textos:
                observacao = ' | '.join(textos)
            
            # Valor final
            valor_final = valor_realizado if valor_realizado is not None else valor_previsto
            if valor_final is None or valor_final == 0:
                continue
            
            # Se não tem vencimento, usar primeiro dia do mês
            if not vencimento:
                vencimento = f"01/{mes_num:02d}/{ano}"
            
            # Categoria baseada na seção
            cat1, cat2, cat3 = '', '', ''
            for chave, cats in SECAO_PARA_CATEGORIA.items():
                if chave in secao_atual.lower():
                    cat1, cat2, cat3 = cats
                    break
            
            # Tipo
            tipo = 'Receita'
            if 'despesa' in secao_atual.lower() or 'equipe' in secao_atual.lower() or 'marketing' in secao_atual.lower() or 'cau' in secao_atual.lower() or 'tecnologia' in secao_atual.lower() or 'montagem' in secao_atual.lower() or 'variav' in secao_atual.lower():
                tipo = 'Despesa'
                if valor_final > 0:
                    valor_final = -valor_final
            
            todos_lancamentos.append({
                'DESCRIÇÃO': descricao,
                'VENCIMENTO': vencimento,
                'DATA': data_pago,
                'VALOR': valor_final,
                'TIPO': tipo,
                '1ª CATEGORIA': cat1,
                '2ª CATEGORIA': cat2,
                '3ª CATEGORIA': cat3,
                'CONCILIADO': 'SIM' if data_pago else 'NÃO',
                'OBSERVAÇÃO': observacao,
            })
    
    if not todos_lancamentos:
        raise ValueError("Nenhum lançamento encontrado. Verifique se a aba selecionada contém dados financeiros.")
    
    df_resultado = pd.DataFrame(todos_lancamentos)
    
    # Remover duplicatas exatas
    df_resultado = df_resultado.drop_duplicates().reset_index(drop=True)
    
    return df_resultado
