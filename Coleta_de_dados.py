import pandas as pd
import winsound
import time

ARQUIVO = r'C:\Temp\DAY_TRADE.xlsm'
ABA = 'Planilha1'

class Gerenciador_excel:
    def __init__(self, caminho, aba):
        self.caminho = caminho
        self.aba= aba
        self.processados = set()
        

    def carregar_excel(self):
        df = pd.read_excel(self.caminho,sheet_name = self.aba, engine='openpyxl')
        return df
    
   
    def filtrar_novos(self, df):
        """Retorna apenas as linhas com Data nova"""
        novos = df[~df["Data"].isin(self.processados)]
        # atualizar o set com as novas datas
        self.processados.update(novos["Data"].unique())
        return novos



if __name__ == '__main__':

    p1 = Gerenciador_excel(ARQUIVO,ABA)
            #print( p1.carregar_excel())
    
    while True:

        df =  p1.carregar_excel()
        df_novos =p1.filtrar_novos(df)

        if not df_novos.empty:
            resultado_compras = df_novos.groupby("Compradora")[["Quantidade"]].sum().rename(columns={"Quantidade": "compras"}) #type: ignore
            resultado_vendas  = df_novos.groupby("Vendedora")[["Quantidade"]].sum().rename(columns={"Quantidade": "vendas"})   #type: ignore

            # juntar por índice
            saldo = resultado_compras.merge(resultado_vendas, left_index=True, right_index=True, how="outer").fillna(0)
            saldo["saldo"] = saldo["compras"] - saldo["vendas"]

            # DEFINO SALDO TOTAL CONSIDERANDO COMPRA E VENDA 
            saldo_compras_alto = saldo[saldo["saldo"]> 200]
            saldo_vendas_alto = saldo[saldo["saldo"] < -100]
            
            # SEPARO OS AGRESSORES POR CATEGORIAS COMPRADOR OU AGRESSOR
            df_comprador = df[df["Agressor"] == "Vendedor"]
            df_vendedor = df[df["Agressor"] == "Comprador"]

            # FILTRANDO APENAS OS PASSIVOS DE VENDA
            resultado_vendas = (
                df_comprador.groupby("Vendedora")[["Quantidade"]]
                .sum()
                .rename(columns={"Quantidade": "vendas"})
            )
           
            # FILTRANDO APENAS OS PASSIVOS DE COMPRA
            resultado_compras = (
                df_vendedor.groupby("Compradora")[["Quantidade"]]
                .sum()
                .rename(columns={"Quantidade": "compras"})
            )
           
            #Define COMPRADOR PASSIVO - ABOSORÇÃO DA VENDA
            passivos_compra = df_comprador.groupby("Compradora")[["Quantidade"]].sum()
            passivos_compra = passivos_compra[passivos_compra["Quantidade"] > 818]
            df_final = saldo_compras_alto.merge(passivos_compra, left_index=True, right_index=True, how="inner")
            for idx, row in df_final.iterrows():
                print(f"Player: {idx} | Saldo: {row['saldo']} | Lotes Comprados: {row['Quantidade']}")

            #Define COMPRADOR PASSIVO - ABOSORÇÃO DA VENDA
            passivos_venda = df_vendedor.groupby("Vendedora")[["Quantidade"]].sum()
            passivos_venda = passivos_venda[passivos_venda["Quantidade"] > 500]
            df_final = saldo_vendas_alto.merge(passivos_venda, left_index=True, right_index=True, how="inner")
            for idx, row in df_final.iterrows():
                print(f"Player: {idx} | Saldo: {row['saldo']} | Lotes Vendidos: {row['Quantidade']}")






           # print (resultado_compras)         
            


        time.sleep(3)