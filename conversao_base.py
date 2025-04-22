import pandas as pd

def padronizar_base(arquivo_cru_path, saida_path):
    # Lista de colunas padrão (modelo 2025)
    colunas_padronizado = [
        'Cd Centro', 'Desc Reg Estrategica', 'Filial', 'Transportador', 'Marca',
        'Tipo NC', 'NM_PAIS', 'Expurgos', 'Desc Ger Mercado', 'Rota', 'Nm Ciclo',
        'Nome da Consultora', 'Desc Ger Vendas', 'Desc Setor Vendas', 'Grupo  (Descrição)',
        'Responsável do Grupo - CRM (Descrição)', 'UF', 'Dc Cidade', 'CN: Bairro (Chave)',
        'Cod Mat', 'Valor', 'Categoria', 'Nível de Risco', 'Status da operação',
        'Solução da NC (Geral)', 'Assunto', 'Resultado', 'Cd Zona Transp', 'Obs',
        'Linha de Separação', 'Data Separação', 'Data Faturamento', 'Data Criação NF',
        'Data de Conclusão', 'Solução segundo nível', 'Origem', 'N° da operação',
        'Nm Pedido', 'Solução Geral', 'Cod Venda', 'Nome Produto', 'Parte Afetada',
        'Cód CN', 'Motivo - CRM (nível 1)', 'Motivo', 'Num Nfe', 'Problema do Produto',
        'Seleção Dimensão_NCProdutos', 'Seleção Dimensão_NCProdutos - CD',
        'Seleção Dimensão_NCProdutos - Cidade', 'Seleção Dimensão_NCProdutos - Filial',
        'Seleção Dimensão_NCProdutos - RE', 'Seleção Dimensão_NCProdutos - UF',
        'Seleção Dimensão_NCProdutos - Zoneamento', 'NC SIM Não', 'Quantidade',
        'Cd Venda Ina', 'Cd Venda Ina-1', 'Chave', 'Chave1', 'Data Criação do Registro',
        'Expurgos1', 'NC Duplicidade', 'NC Expurgo', 'Nm Transportadora', 'NM_PAIS.1',
        'Pedido Conquista'
    ]

    # Mapeamento das colunas do arquivo cru para o padrão
    mapeamento = {
        'Cd Centro': None,
        'Desc Reg Estrategica': None,
        'Filial': 'filial',
        'Transportador': 'nome_transportadora',
        'Marca': 'marca_canal',
        'Tipo NC': 'nc_type',
        'NM_PAIS': None,
        'Expurgos': None,
        'Desc Ger Mercado': 'gm',
        'Rota': 'rota',
        'Nm Ciclo': 'nm_ciclo',
        'Nome da Consultora': None,
        'Desc Ger Vendas': 're',
        'Desc Setor Vendas': 'cd_setor',
        'Grupo  (Descrição)': None,
        'Responsável do Grupo - CRM (Descrição)': 'owner_email',
        'UF': 'estado',
        'Dc Cidade': 'cidade',
        'CN: Bairro (Chave)': None,
        'Cod Mat': 'nat_materialid',
        'Valor': None,
        'Categoria': 'categoria',
        'Nível de Risco': None,
        'Status da operação': 'order_status',
        'Solução da NC (Geral)': 'nat_solution',
        'Assunto': 'reason',
        'Resultado': None,
        'Cd Zona Transp': 'cd_zona_transporte',
        'Obs': 'observacao_falta',
        'Linha de Separação': None,
        'Data Separação': 'data_hr_separacao',
        'Data Faturamento': 'data_hr_faturamento',
        'Data Criação NF': 'data_hr_envio_linha',
        'Data de Conclusão': 'data_conclusao',
        'Solução segundo nível': 'nat_second_solution_name',
        'Origem': 'caseorigin_entity_name',
        'N° da operação': 'ticket_number',
        'Nm Pedido': 'nm_pedido',
        'Solução Geral': None,
        'Cod Venda': 'nat_productsellcode',
        'Nome Produto': 'product_description',
        'Parte Afetada': None,
        'Cód CN': 'cd_cn',
        'Motivo - CRM (nível 1)': None,
        'Motivo': None,
        'Num Nfe': 'nat_invoice_number',
        'Problema do Produto': None,
        'Seleção Dimensão_NCProdutos': None,
        'Seleção Dimensão_NCProdutos - CD': None,
        'Seleção Dimensão_NCProdutos - Cidade': None,
        'Seleção Dimensão_NCProdutos - Filial': None,
        'Seleção Dimensão_NCProdutos - RE': None,
        'Seleção Dimensão_NCProdutos - UF': None,
        'Seleção Dimensão_NCProdutos - Zoneamento': None,
        'NC SIM Não': None,
        'Quantidade': 'nat_quantity',
        'Cd Venda Ina': None,
        'Cd Venda Ina-1': None,
        'Chave': None,
        'Chave1': None,
        'Data Criação do Registro': 'current_timestamp()',
        'Expurgos1': None,
        'NC Duplicidade': None,
        'NC Expurgo': None,
        'Nm Transportadora': 'nome_transportadora',
        'NM_PAIS.1': None,
        'Pedido Conquista': 'natura_order',
    }

    # Ler arquivo cru
    df_cru = pd.read_excel(arquivo_cru_path)

    # Criar DataFrame final
    df_final = pd.DataFrame()
    for coluna in colunas_padronizado:
        if mapeamento.get(coluna) in df_cru.columns:
            df_final[coluna] = df_cru[mapeamento[coluna]]
        else:
            df_final[coluna] = None

    # Salvar CSV
    df_final.to_csv(saida_path, sep=";", index=False, encoding="utf-8-sig")
    print(f"✅ Arquivo padronizado salvo com sucesso em: {saida_path}")


# === CAMINHOS DO ARQUIVO ===
entrada = r"C:\Users\atend\OneDrive\Documentos\GitHub\ncfaltas\raw\NC Falta JB 16.04.xlsx"
saida = r"C:\Users\atend\OneDrive\Documentos\GitHub\ncfaltas\padronizado\arquivo_padronizado.csv"

# Executar
padronizar_base(entrada, saida)
