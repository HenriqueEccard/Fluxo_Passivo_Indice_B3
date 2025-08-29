import pandas as pd
import winsound
import time

ARQUIVO = r'C:\Temp\DAY_TRADE.xlsm'
ABA = 'Planilha1'

class Gerenciador_excel:
    def __init__(self, caminho, aba):
        self.caminho = caminho
        self.aba = aba
        self.processados = set()
        
    def carregar_excel(self):
        df = pd.read_excel(self.caminho, sheet_name=self.aba, engine='openpyxl')
        return df
    
    def filtrar_novos(self, df):
        """Retorna apenas as linhas com Data nova"""
        novos = df[~df["Data"].isin(self.processados)]
        self.processados.update(novos["Data"].unique())
        return novos


if __name__ == '__main__':

    p1 = Gerenciador_excel(ARQUIVO, ABA)
    
    while True:
        df = p1.carregar_excel()
        df_novos = p1.filtrar_novos(df)

        if not df_novos.empty:
            # 1. Soma dos players compradores (agressor comprador)
            df_comprador = df_novos[df_novos["Agressor"] == "Comprador"]
            resultado_compras = (
                df_comprador.groupby("Compradora")[["Quantidade"]]
                .sum()
                .rename(columns={"Quantidade": "compras"})
            )

            # 2. Soma dos players vendedores (agressor vendedor)
            df_vendedor = df_novos[df_novos["Agressor"] == "Vendedor"]
            resultado_vendas = (
                df_vendedor.groupby("Vendedora")[["Quantidade"]]
                .sum()
                .rename(columns={"Quantidade": "vendas"})
            )

            # 3. Saldo geral (compras - vendas)
            saldo = resultado_compras.merge(resultado_vendas, 
                                            left_index=True, 
                                            right_index=True, 
                                            how="outer").fillna(0)
            saldo["saldo"] = saldo["compras"] - saldo["vendas"]

            print("\n==== SALDO ATUAL ====")
            print(saldo)

            # 4. Condição: agressor comprador > 200 e saldo > 200
            cond1 = saldo.merge(resultado_compras, left_index=True, right_index=True, how="inner")
            cond1_filtrado = cond1[(cond1["compras_y"] > 200) & (cond1["saldo"] > 200)]

            for idx, row in cond1_filtrado.iterrows():
                winsound.Beep(700, 500)
                print(f"[PASSIVO VENDA] Player {idx} | Total Comprado={row['compras_y']} | Saldo={row['saldo']}")

            # 5. Condição: agressor vendedor > 200 e saldo > 200 (invertendo sinal se saldo negativo)
            cond2 = saldo.merge(resultado_vendas, left_index=True, right_index=True, how="inner")
            cond2["saldo_corrigido"] = cond2["saldo"].apply(lambda x: abs(x))  # inverte negativo
            cond2_filtrado = cond2[(cond2["vendas_y"] > 200) & (cond2["saldo_corrigido"] > 200)]

            for idx, row in cond2_filtrado.iterrows():
                winsound.Beep(1000, 500)
                print(f"[PASSIVO COMPRA] Player {idx} | Total Vendido={row['vendas_y']} | Saldo Corrigido={row['saldo_corrigido']}")

        # aguarda 3 segundos antes da próxima leitura
        time.sleep(3)
