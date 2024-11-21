import pandas as pd

# Carregar a planilha
df = pd.read_excel('./src/Orientadores entrada.xlsx')


# Criar um dicionário para armazenar os resultados
relatorio = {}

# # Iterar sobre cada coluna para calcular os quantitativos absolutos e relativos
for coluna in df.columns:
    # Contagem absoluta dos valores na coluna
    contagem_absoluta = df[coluna].value_counts(dropna=False)
    
    # Contagem relativa excluindo valores em branco
    contagem_relativa = df[coluna].value_counts(normalize=True) * 100
    
    # Contagem relativa incluindo valores em branco
    total_respostas_incluindo_brancos = len(df[coluna])
    contagem_relativa_incluindo_brancos = (contagem_absoluta / total_respostas_incluindo_brancos) * 100
    
    # Combinar contagens absoluta e relativa em um DataFrame
    relatorio_coluna = pd.DataFrame({
        'Contagem Absoluta': contagem_absoluta,
        'Contagem Relativa (%)': contagem_relativa,
        'Contagem Relativa com Em Brancos (%)': contagem_relativa_incluindo_brancos
    })
    
    # Substituir valores em branco pelo texto "Em Branco"
    relatorio_coluna.index = relatorio_coluna.index.fillna('Em Branco')
    
    # Adicionar o relatório da coluna ao dicionário
    relatorio[coluna] = relatorio_coluna

# Salvar o relatório em um arquivo Excel
with pd.ExcelWriter('./Saida/Frequência Orientadores.xlsx') as writer:
    for coluna, relatorio_coluna in relatorio.items():
        relatorio_coluna.to_excel(writer, sheet_name=coluna.split(" ")[0])
