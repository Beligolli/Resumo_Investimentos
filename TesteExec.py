# %%
import pandas as pd

base_dash_df = pd.read_excel(
    r'base_dash.xlsx', index_col=[0])
infos_df = pd.read_excel(
    r'infos.xlsx', index_col=[0])
# display(base_dash_df)
# display(infos_df)


# %% [markdown]
# Concatenar com novas infos

# %%
bases_df = [base_dash_df, infos_df]
base_dash_df = pd.concat(bases_df,)
# print(base_dash_df)
base_dash_df.to_excel(r'base_dash.xlsx')


# %%
base_dash_df.plot(x="Data", figsize=(20, 10))
# display(base_dash_df)

# %%
from tkinter import *
from tkinter import messagebox
janela = Tk()
messagebox.showinfo('Status', 'Processo Concluído.')
janela.destroy


