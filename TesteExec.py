import pandas as pd
import win32com.client as win32

# Start an instance of Excel
xlapp = win32.DispatchEx("Excel.Application")

# Open the workbook in said instance of Excel
wb = xlapp.workbooks.open(r'C:\Users\belig\OneDrive\Python\MeuProjeto\Projetos\dash_investimentos\Resumo_Investimentos\infos.xlsx')

# Optional, e.g. if you want to debug
xlapp.Visible = True

# Refresh all data connections.
wb.RefreshAll()
wb.Save()

# Quit
xlapp.Quit()

base_dash_df = pd.read_excel(
    r'C:\Users\belig\OneDrive\Python\MeuProjeto\Projetos\dash_investimentos\Resumo_Investimentos\base_dash.xlsx', index_col=[0])
infos_df = pd.read_excel(
    r'C:\Users\belig\OneDrive\Python\MeuProjeto\Projetos\dash_investimentos\Resumo_Investimentos\infos.xlsx', index_col=[0])

bases_df = [base_dash_df, infos_df]
base_dash_df = pd.concat(bases_df, )
# base_dash_df.iloc[-1]['Dow Jones'] = base_dash_df.iloc[-1]['Dow Jones'] * cotacao_dolar
# print(base_dash_df.iloc[-1]['Dow Jones'])
print(base_dash_df)
base_dash_df.to_excel(r'C:\Users\belig\OneDrive\Python\MeuProjeto\Projetos\dash_investimentos\Resumo_Investimentos\base_dash.xlsx')

posicao_detalhada_df = pd.read_excel(
    r'C:\Users\belig\OneDrive\Python\MeuProjeto\Projetos\dash_investimentos\Resumo_Investimentos\PosicaoDetalhada.xlsx', index_col=[0])
posicao = posicao_detalhada_df.columns[4]
data_posicao_detalhada = posicao[17:27]
print(data_posicao_detalhada)