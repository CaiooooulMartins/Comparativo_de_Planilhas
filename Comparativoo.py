import pandas as pd
from datetime import datetime



def comparar_planilhas(caminho_antigo, caminho_novo, colunas_comparar, chave_primaria='informar chave primaria'):

    
    # Carregar as planilhas preenchendo valores nulos
    try:
        antigo = pd.read_excel(caminho_antigo).fillna('VAZIO')
        novo = pd.read_excel(caminho_novo).fillna('VAZIO')
    except Exception as e:
        print(f"Erro ao carregar planilhas: {e}")
        return

    # Verificar se colunas existem
    print("\nVerificando estrutura das planilhas...")
    for col in colunas_comparar + [chave_primaria]:
        if col not in antigo.columns:
            print(f"Aviso: Coluna '{col}' não encontrada na planilha antiga")
        if col not in novo.columns:
            print(f"Aviso: Coluna '{col}' não encontrada na planilha nova")

    # Identificar linhas adicionadas/removidas
    print("\nIdentificando linhas adicionadas/removidas...")
    comparacao = pd.merge(antigo, novo, how='outer', indicator=True)
    linhas_removidas = comparacao[comparacao['_merge'] == 'left_only'].drop(columns=['_merge'])
    linhas_adicionadas = comparacao[comparacao['_merge'] == 'right_only'].drop(columns=['_merge'])

    # Comparar valores nas linhas existentes em ambas
    print("Comparando valores nas linhas existentes...")
    df_comp = pd.merge(antigo, novo, on=chave_primaria, suffixes=('_ANTIGO', '_NOVO'))
    
    # Identificar alterações
    alteracoes = []
    for coluna in colunas_comparar:
        col_antiga = f"{coluna}_ANTIGO"
        col_nova = f"{coluna}_NOVO"
        
        if col_antiga in df_comp.columns and col_nova in df_comp.columns:
            df_comp[f'{coluna}_ALTERADO'] = df_comp[col_antiga] != df_comp[col_nova]
            alteracoes.append(f'{coluna}_ALTERADO')
    
    linhas_alteradas = df_comp[df_comp[alteracoes].any(axis=1)] if alteracoes else pd.DataFrame()

    # Preparar relatório detalhado
    print("Preparando relatório...")
    if not linhas_alteradas.empty:
        for coluna in colunas_comparar:
            if f"{coluna}_ANTIGO" in linhas_alteradas.columns and f"{coluna}_NOVO" in linhas_alteradas.columns:
                linhas_alteradas[f'{coluna}_MUDANCA'] = linhas_alteradas.apply(
                    lambda x: f"{x[f'{coluna}_ANTIGO']} → {x[f'{coluna}_NOVO']}" 
                    if x[f'{coluna}_ANTIGO'] != x[f'{coluna}_NOVO'] else '', axis=1)

    # Gerar arquivo de relatório
    nome_relatorio = f"relatorio_alteracoes_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    try:
        with pd.ExcelWriter(nome_relatorio) as writer:
            # Adicionar metadados
            pd.DataFrame({
                'Informação': ['Data de execução', 'Planilha Antiga', 'Planilha Nova', 'Total Removidos', 'Total Adicionados', 'Total Alterados'],
                'Valor': [
                    datetime.now().strftime('%d/%m/%Y %H:%M'),
                    caminho_antigo.split('\\')[-1],
                    caminho_novo.split('\\')[-1],
                    len(linhas_removidas),
                    len(linhas_adicionadas),
                    len(linhas_alteradas)
                ]
            }).to_excel(writer, sheet_name='Resumo das alterações', index=False)
            
            # Adicionar abas com resultados
            if not linhas_removidas.empty:
                linhas_removidas.to_excel(writer, sheet_name='Removidas', index=False)
            if not linhas_adicionadas.empty:
                linhas_adicionadas.to_excel(writer, sheet_name='Adicionadas', index=False)
            if not linhas_alteradas.empty:
                # Ordenar por chave primária e selecionar colunas relevantes
                cols_relatorio = [chave_primaria] + \
                               [c for c in linhas_alteradas.columns if '_MUDANCA' in c] + \
                               [c for c in linhas_alteradas.columns if '_ANTIGO' in c or '_NOVO' in c]
                
                linhas_alteradas[cols_relatorio].sort_values(chave_primaria) \
                    .to_excel(writer, sheet_name='Alteradas', index=False)
        
        print(f"\nRelatório gerado com sucesso: {nome_relatorio}")
        print(f"Resumo: {len(linhas_removidas)} removidos | {len(linhas_adicionadas)} adicionados | {len(linhas_alteradas)} alterados")
        
    except Exception as e:
        print(f"Erro ao gerar relatório: {e}")

# Configuração
if __name__ == "__main__":
    # Defina aqui os parâmetros
    CAMINHO_ANTIGO = "INFORMAR O CAMINHO DO ARQUIVO ANTIGO"
    CAMINHO_NOVO =  "INFORMAR O CAMINHO DO ARQUIVO NOVO"
    COLUNAS_COMPARAR = ['COLUNA 1', 'COLUNA 2', 'COLUNA 3', 'COLUNA 3', 'COLUNA 4', 'COLUNA 5']
    CHAVE_PRIMARIA = 'Informar o nome da coluna chave primária (ex: Numero de produto,nome do item,etc)'
    
    # Executar comparação
    comparar_planilhas(CAMINHO_ANTIGO, CAMINHO_NOVO, COLUNAS_COMPARAR, CHAVE_PRIMARIA)