# 🚚 **Comparador de Planilhas Excel**

## 📜 Descrição

O **Comparador de Planilhas Excel** é uma ferramenta desenvolvida em **Python** utilizando a biblioteca **Pandas** para análise de dados. O objetivo é comparar duas versões de uma planilha (antiga e nova) e gerar um relatório detalhado com todas as alterações encontradas.

## 🛠️ **Estrutura do Projeto**

### 1. 📚 **Bibliotecas Utilizadas**

- **pandas**: Para manipulação e análise dos dados
- **openpyxl**: Para leitura/escrita de arquivos Excel
- **datetime**: Para registro do momento da execução

### 2. ⚙️ **Configuração Inicial**

O script foi desenvolvido para:
- Tratar automaticamente valores nulos
- Verificar a existência das colunas especificadas
- Gerar relatório com timestamp no nome

---

## 🏷️ **Funcionalidades Principais**

### 1. 🔍 **Identificação de Diferenças**
- Detecta linhas **adicionadas** na nova versão
- Identifica linhas **removidas** da versão antiga
- Compara valores nas linhas existentes em ambas versões

### 2. 📊 **Geração de Relatório**
Relatório em Excel contendo:
- **Resumo das alterações**: Totais e metadados
- **Removidas**: Linhas exclusivas da versão antiga
- **Adicionadas**: Linhas exclusivas da nova versão  
- **Alteradas**: Linhas com modificações nos valores

### 3. ✅ **Validações Automáticas**
- Verifica se colunas especificadas existem em ambas planilhas
- Confere se a chave primária contém valores únicos
- Trata valores nulos para evitar falsas diferenças

---

## 🏁 **Passo a Passo do Funcionamento**

1. **Configuração**: Editar no final do script:
   - Caminhos das planilhas
   - Colunas para comparação
   - Coluna chave primária

2. **Execução**: Rodar o script via terminal:
   ```bash
   python Comparativoo.py
   ## 📤 **Saída do Relatório**

3. **Saida**: O script gera automaticamente um arquivo Excel contendo:

✔ **Listagem completa** de todas as diferenças encontradas  
✔ **Comparação lado a lado** dos valores antigos e novos  
✔ **Destaque automático** para a opção mais vantajosa (menor valor)

---

## 📋 **Exemplo de Relatório Gerado**

### 📊 **Resumo das Alterações**
| Ícone | Item | Valor |
|-------|------|-------|
| 📅 | Data de execução | 15/07/2023 14:30 |
| 📂 | Planilha Antiga | dados_jan.xlsx |
| 📂 | Planilha Nova | dados_fev.xlsx |
| ➖ | Registros Removidos | 12 |
| ➕ | Registros Adicionados | 18 |
| 🔄 | Registros Alterados | 7 |



## ⚠️ **Limitações Técnicas**

### 📦 **Tamanho de Arquivos**
- Planilhas muito extensas podem requerer:
  - Maior capacidade de memória RAM
  - Tempo adicional de processamento

### 🔢 **Tipos de Dados**
- **Não analisa**:
  - Fórmulas de células
  - Formatação visual
  - Comentários e anotações
- **Limitado a**:
  - Valores numéricos
  - Textos simples
  - Datas formatadas

### 🎯 **Requisitos Obrigatórios**
- Coluna de identificação única:
  - Deve conter valores não repetidos
  - Pode ser: ID, Código, CPF, etc.
  - Essencial para comparação precisa

### 🔮 Futuras Funcionalidades
- [ ] Interface gráfica intuitiva

