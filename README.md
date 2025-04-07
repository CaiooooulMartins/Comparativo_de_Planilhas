📝 Descrição
Este script Python compara duas versões de uma planilha Excel (uma versão antiga e uma nova) e gera um relatório detalhado com todas as alterações encontradas, incluindo:

Linhas removidas

Linhas adicionadas

Linhas com valores alterados

📋 Pré-requisitos
Python 3.6 ou superior

Bibliotecas necessárias:

bash
Copy
pip install pandas openpyxl
🛠 Configuração
Edite as variáveis no final do arquivo Comparativoo.py:

python
Copy
CAMINHO_ANTIGO = "caminho/para/planilha_antiga.xlsx"
CAMINHO_NOVO = "caminho/para/planilha_nova.xlsx"
COLUNAS_COMPARAR = ['coluna1', 'coluna2', 'coluna3']  # Colunas a comparar
CHAVE_PRIMARIA = 'coluna_chave'  # Coluna com identificador único
Salve o arquivo após as alterações

▶️ Como Executar
bash
Copy
python Comparativoo.py
📊 Saída
O script gera um arquivo Excel chamado relatorio_alteracoes_AAAAMMDD_HHMM.xlsx com três abas:

Resumo das alterações: Metadados e totais de mudanças

Removidas: Linhas presentes apenas na planilha antiga

Adicionadas: Linhas presentes apenas na planilha nova

Alteradas: Linhas com diferenças nos valores

🔍 Detalhes Técnicos
Valores nulos são tratados como 'VAZIO' para comparação

O relatório mostra mudanças no formato "valor_antigo → valor_novo"

As linhas alteradas são ordenadas pela chave primária

O script verifica se as colunas especificadas existem em ambas planilhas

💡 Melhorias Futuras
Adicionar suporte para múltiplas abas

Incluir formatação condicional no Excel de saída

Adicionar opção para comparação case-insensitive

Implementar comparação de tipos de dados

⚠️ Limitações
Requer que a chave primária seja única em ambas planilhas

Planilhas muito grandes podem exigir otimizações de memória

Não compara formatação ou fórmulas, apenas valores
