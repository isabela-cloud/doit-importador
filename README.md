# DOit Importador

Conversor de planilhas de clientes para o padrão de importação do DOit. Recebe dados de diversos sistemas (Navis, Sienge, ClickUp, Conta Azul, etc.) e gera arquivos prontos para importação com abas auxiliares.

## Estrutura de Pastas

```
doit-importador/
├── data/                  # Planilhas de exemplo por sistema (navis-*, sienge-*, clickup-*, etc.)
├── input/                 # Arquivos recebidos dos clientes para conversão
├── models/                # Modelos padrão DOit (modelo_contatos.xlsx, modelo_financeiro.xlsx, etc.)
├── mappings/              # Mapeamentos e listas auxiliares (.txt)
├── scripts/               # Código principal (conversor.py, parser_navis.py, parser_bpo.py, abas_auxiliares.py)
├── dashboards/            # Interface Streamlit (app_conversor.py)
├── output/                # Arquivos convertidos gerados
├── docs/                  # Documentação
├── requirements.txt       # Dependências Python
└── README.md
```

## Dependências

```bash
pip install -r requirements.txt
```

## Como Executar

```bash
streamlit run dashboards/app_conversor.py
```

## Sistemas Suportados

| Sistema | Tipos de Dados |
|---------|---------------|
| DOit Coleta (Planilha Padrão) | Contatos, Projetos, Financeiro |
| Navis | Financeiro CC, Previsão, Projetos, Clientes, Fornecedores, Contatos, Horas |
| Sienge | Contatos, Projetos, Financeiro, Horas |
| ClickUp | Contatos, Projetos, Horas, Financeiro |
| Conta Azul | Contatos, Financeiro |
| Trello | Contatos, Projetos, Horas, Financeiro |
| BPO (Financeiro Horizontal) | Financeiro |
| Excel Desestruturado | Todos |

## Funcionalidades

- Detecção automática de cabeçalho e tipo de relatório
- Mapeamento inteligente de colunas (automático + manual)
- Formatação de CPF/CNPJ, telefones, CEP, datas
- Geração de abas auxiliares: Dados Bancários, Plano de Contas, Pendências
- Validação e alertas de padronização
- Vínculo de IDs entre cadastros e projetos
- Suporte a múltiplos estilos de caixa (Primeira Maiúscula, MAIÚSCULA, Original)

## Padrão de Nomes dos Arquivos de Exemplo

Os arquivos na pasta `data/` seguem o padrão: `SISTEMA-tipo.xlsx`

Exemplos:
- `navis-financeiro-cc.xlsx`
- `navis-clientes.xlsx`
- `sienge-exemplo.xlsx`
- `clickup-exemplo.xlsx`
- `contaazul-financeiro.xlsx`
- `omie-financeiro.xlsx`
