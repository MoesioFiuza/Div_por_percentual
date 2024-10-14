import pandas as pd
import os

planilha_matriz = r'C:\Users\moesios\Desktop\HORA\1- EDU_COL_unificada.xlsx'
diretorio_saida = r'T:\ANDAMENTO\2313-RS-POA-SMA-COLDA\03-PRODUTOS\03.3-COLETA DADOS\03.3.2-FINAL\PRODUTO 14 - REV00\Matriz OD - VIAGENS\OD VIAGENS - HORA A HORA\EDUIND'

if not os.path.exists(diretorio_saida):
    os.makedirs(diretorio_saida)

df_matriz = pd.read_excel(planilha_matriz, sheet_name='Sheet1')

percentuais = [
    0.028233, 0.000000, 0.000000, 0.000000, 0.000000, 0.056465,
    4.009034, 21.569735, 3.105590, 1.016375, 0.762281, 4.206663,
    21.456804, 10.135517, 0.875212, 0.988142, 2.399774, 14.596273,
    7.509881, 1.863354, 0.818746, 1.637493, 2.625635, 0.338792
]

soma_total = 0

for i, percentual in enumerate(percentuais):
    df_modificado = df_matriz.copy()
    for col in df_modificado.columns:
        if df_modificado[col].dtype in ['int64', 'float64']:
            df_modificado[col] = df_modificado[col] * (percentual / 100)
    soma_total += df_modificado.select_dtypes(include=['int64', 'float64']).sum().sum()
    arquivo_saida = os.path.join(diretorio_saida, f'EDUCOL_HORA_{i}.xlsx')
    df_modificado.to_excel(arquivo_saida, index=False)

print(f"Soma total dos valores das 24 planilhas modificadas: {soma_total}")
