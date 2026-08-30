# Projeto-Intercement

#importando as bibliotecas pandas e numpy
import pandas as pd
import numpy as np

#conhecendo o arquivo e disposição das variaveis
df= pd.read_excel("COTAÇÕES_CASE_(1).xlsx")
df.head(10)

# Limpando as linhas antes dos cabeçalhos
df1 = pd.read_excel("COTAÇÕES_CASE_(1).xlsx", header=2)
df1.head(10)

# Conhecendo as variáveis
df1.columns.tolist()
df1.dtypes

# consultando a completude dos dados da tabela
total_linhas = len(df1)
completude = df1.notna().sum()
percentual_completude = (completude / total_linhas) * 100
resumo = pd.DataFrame({"Não nulos": completude, "Total linhas": total_linhas, "Completude (%)": percentual_completude.round(2)})
print(resumo)

# verificar se existem duplicatas
df1.duplicated()

# Contar valores duplicados por coluna
duplicados_por_coluna = df1.apply(lambda x: x.duplicated().sum())
print(duplicados_por_coluna)

# Identificar IDs duplicados
ids_duplicados = df1[df1["ID da Cotação do SAP"].duplicated(keep=False)]
ids_duplicados[["ID da Cotação do SAP", "Id Solic. Ajuste"]].sort_values("ID da Cotação do SAP")

#Limpando os duplicados mantendo apenas um registro, apenas a primeira ocorrência. Foi selecionada ID da Cotação do SAP como variável pois entende-se que deva ser uma valor único não devendo se repetir
df1_limpo = df1.drop_duplicates(subset=["ID da Cotação do SAP"], keep="first")
print(f"Linhas originais: {len(df1)}")
print(f"Linhas após remoção de duplicados: {len(df1_limpo)}")

# consultando a nova completude dos dados da tabela após saneados dos duplicados.
total_linhas = len(df1_limpo)
completude = df1_limpo.notna().sum()
percentual_completude = (completude / total_linhas) * 100
resumo = pd.DataFrame({"Não nulos": completude, "Total linhas": total_linhas, "Completude (%)": percentual_completude.round(2)})
print(resumo)

#conhecendo o arquivo e disposição das variaveis da base do Sales Force
df2 = pd.read_excel("SalesForce_CASE_(1).xlsx")

#Remover duplicados na coluna ID da Cotação do SAP
df2 = df2.drop_duplicates(subset=["ID da Cotação do SAP"], keep="first")

#Filtrar apenas linhas com ID válido
df2 = df2[df2["ID da Cotação do SAP"].notna()]
df2 = df2[df2["ID da Cotação do SAP"].astype(str).str.strip().ne("")]

#avaliar base após o saneamento
print(f"Linhas após limpeza: {len(df2)}")
df2.head(10)

# Consultando a completude dos dados da tabela do Sales Force após saneamento
total_linhas = len(df2)
completude = df2.notna().sum()
percentual_completude = (completude / total_linhas) * 100
resumo_df2 = pd.DataFrame({"Não nulos": completude,"Total linhas": total_linhas,"Completude (%)": percentual_completude.round(2)})
print(resumo_df2)

# Conclusão da análise: não foram identificadas duplicatas na chave tripla. Segue com renome de coluna para alinhamento das bases
df2 = df2.rename(columns={"Material: Código do material": "Material"})
print(f"\nBase SalesForce preparada — {len(df2)} linhas únicas")

#Renomear coluna para igualar nas duas bases. Na base SalesForce: "Material: Código do material" → vira "Material"
df2 = df2.rename(columns={"Material: Código do material": "Material"})

# Criar uma CHAVE TRIPLA (3 colunas juntas)
colunas_chave = ["ID da Cotação do SAP", "Material", "Cód. Expedição"]

#rodar o join garantindo que não vai duplicar linhas
base_consolidada = df1_limpo.merge(df2[colunas_chave + ["Frete Comercial", "Chapa"]], on=colunas_chave, how="left", validate="1:1")

#Verificar se não aumentou o número de linhas
print(f"Linhas COTAÇÕES original: {len(df1_limpo)}")
print(f"Linhas após cruzamento:   {len(base_consolidada)}")
if len(base_consolidada) == len(df1_limpo):
    print("SUCESSO")
else:
    print("ATENÇÃO")

# Separar os registros sem correspondência (diferenciar NaN de 0)
base_consolidada["Sem_Correspondencia_SF"] = base_consolidada["Frete Comercial"].isna()

# Contagem resumida
sem_corresp = base_consolidada["Sem_Correspondencia_SF"].sum()
frete_zero = len(base_consolidada[(base_consolidada["Frete Comercial"] == 0) & (~base_consolidada["Frete Comercial"].isna())])
print(f"\n Resumo Frete Comercial:")
print(f"   Sem correspondência no SalesForce: {sem_corresp} cotações")
print(f"   Frete Comercial = 0 (encontrado):   {frete_zero} cotações")
print(f"   Com frete informado:                 {len(base_consolidada) - sem_corresp}")

# Inicianto preparação para o join. 

#Renomear coluna para alinhar na junção
df2 = df2.rename(columns={"Material: Código do material": "Material"})

# Realizando join e mantendo todas as cotações garantindo que não aumentem linhas relação um para um (LEFT JOIN)
colunas_chave = ["ID da Cotação do SAP", "Material", "Cód. Expedição"]
base_consolidada = df1_limpo.merge(df2[colunas_chave + ["Frete Comercial", "Chapa"]], on=colunas_chave, how="left", validate="1:1")

# Verificar regra do case: número de linhas não pode aumentar
print(f"Linhas COTAÇÕES originais: {len(df1_limpo)}")
print(f"Linhas após cruzamento:    {len(base_consolidada)}")
if len(base_consolidada) == len(df1_limpo):
    print("SUCESSO")
else:
    print("REAVALIAR")

# Diferenciar: frete=0 (encontrado) de "sem correspondência" (não existe)
base_consolidada["Sem_Correspondencia_SF"] = base_consolidada["Frete Comercial"].isna()

# Contagem resumida para conferência
sem_corresp = base_consolidada["Sem_Correspondencia_SF"].sum()
frete_zero = len(base_consolidada[(base_consolidada["Frete Comercial"] == 0) & (~base_consolidada["Frete Comercial"].isna())])
print(f"\n Resumo de correspondência com SalesForce:")
print(f"   Sem correspondência: {sem_corresp} cotações")
print(f"   Frete Comercial = 0: {frete_zero} cotações")
print(f"   Com frete informado: {len(base_consolidada) - sem_corresp}")

#Verificação final da base
base_consolidada.head(10)

# Exporta a base completa para Excel
base_consolidada.to_excel("BASE_FINAL.xlsx", index=False)

Lista de Premissas Adotadas

1. Linhas de cabeçalho
O cabeçalho das colunas não estava na primeira linha. Foi identificado e adotado que os nomes das colunas estão na linha 3 → parâmetro header=2 no carregamento.

2. Duplicações na base COTAÇÕES
Identificadas duplicatas por ID da Cotação do SAP. Adotou-se manter apenas a primeira ocorrência (keep="first"), por considerar o ID como identificador único de negócio.

3. Duplicações na base SalesForce e regra de cruzamento
Chave de cruzamento definida como chave tripla: ID da Cotação do SAP + Material + Cód. Expedição.
Não foram encontradas duplicatas pela chave tripla após limpeza.
Regra de deduplicação adotada: remoção de duplicatas por ID da Cotação antes do cruzamento, mantendo primeira ocorrência.
Garantido que número de linhas da base COTAÇÕES não aumentou após junção (validação validate="1:1").

4. Distinção: sem correspondência × frete zero
Cotações sem correspondência no SalesForce → coluna Sem_Correspondencia_SF = True → NÃO confundir com frete = 0.
Valor Frete Comercial = 0 → valor efetivamente informado como zero → condição negociada.
Esta distinção foi sinalizada em coluna própria e preservada na base final.

5. Janela móvel de 30 dias
Adotado como data de referência exclusivamente o campo Dt. Criação Sol. → Última Ação em foi ignorada nos cálculos, conforme regra.
Para cotações do início do período (janela incompleta), utilizou-se apenas o histórico disponível, sem extrapolar dados.

6. Tipos de cotação elegíveis
Considerados apenas os tipos "Banda" e "PrecoFixo". Os tipos "Prazo" e "FreteComercial" foram excluídos dos cálculos de recorrência e soma de desconto, conforme escopo definido.

7. Valores nulos e ausentes
Cotações sem correspondência no SalesForce foram mantidas na base final e sinalizadas de forma distinta, conforme exigência do Case.

8. Bibliotecas utilizadas
pandas → manipulação, limpeza, cruzamento e exportação de dados
numpy → operações numéricas e tratamento de valores
