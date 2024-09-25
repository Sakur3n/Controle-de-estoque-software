import pandas as pd
from tkinter import *

class Arquivo():
    def limpa_tela(self):
        self.add_product.delete(0, END)

    def limpa_reduzicao(self):
        self.reduzir_produto.delete(0, END)

    def limpa_manual(self):
        self.add_peso.delete(0, END)

    def limpa_manuals(self):
        self.add_pesos.delete(0, END)
        

    def lendo(self):        

        cod_inteiro = self.add_product.get()        
        self.seis_primeiros = int(cod_inteiro[1:7])
        self.seis_ultimos = int(cod_inteiro[7:13])
        self.limpa_tela()

    def somando(self):
        df = pd.read_excel("EstoqueMatriz.xlsx")        
        self.lendo()

        cod_produto = int(self.seis_primeiros)
        dt = df.loc[df["Código"] == cod_produto]
        linha = dt.index[0]

        print('\n', dt, '\n', self.seis_primeiros, '\n', df.at[linha, "Descrição"],'\n')

        if not dt.empty:
            linha = dt.index[0]

            peso_anterior = float(df.at[linha, "estoque atual"])
            preco_venda = float(df.at[linha,"Preço Unitário de Venda"])            
            seis_ult = float(self.seis_ultimos)

            peso_pelo_preco = seis_ult / preco_venda
            peso_etiqueta = float(peso_pelo_preco / 1000)
            peso_etiqueta_arrendodado = round(peso_etiqueta, 3)
            novo_peso = peso_anterior + peso_etiqueta_arrendodado
            
            print("Peso anterior: ",peso_anterior)
            print("Peso da etiqueta: ",peso_etiqueta_arrendodado)
            print("Novo peso na planilha: ",novo_peso)
            
            print('\n -----/-------/--------/---------/--- ')
            print('-----/-------/--------/---------/--- ')

            df.at[linha, "estoque atual"] = novo_peso
            df.to_excel("EstoqueMatriz.xlsx", index = False)

            with open("Balanço.csv", "a", newline="", encoding="utf-8") as arquivo:
                linhaa = dt.at[linha,"Descrição"]
                arquivo.write(f"{linhaa}, {peso_etiqueta_arrendodado}Kg \n ")
                arquivo.closed

        else:
            print("Linha não encontrada") 


    def reading (self):        
        self.cod_inteiro = self.reduzir_produto.get()        
        self.seis_primeiros = self.cod_inteiro[1:7]        
        self.seis_ultimos = self.cod_inteiro[7:13]
        self.limpa_reduzicao()

    def subtraindo (self):
        
        df = pd.read_excel("EstoqueMatriz.xlsx")        
        self.reading()

        cod_produto = int(self.seis_primeiros)
        dt = df.loc[df["Código"] == cod_produto]
        print(dt)

        if not dt.empty:
            linha = dt.index[0]

            peso_anterior = float(df.at[linha, "estoque atual"])
            preco_venda = float(df.at[linha,"Preço Unitário de Venda"])            
            seis_ult = float(self.seis_ultimos)

            peso_pelo_preco = seis_ult / preco_venda
            peso_etiqueta = float(peso_pelo_preco / 1000)
            peso_etiqueta_arrendodado = round(peso_etiqueta, 3)
            novo_peso = peso_anterior - peso_etiqueta_arrendodado
            
            print(peso_etiqueta_arrendodado)
            print(round(novo_peso, 3))
            df.at[linha, "estoque atual"] = novo_peso
            print('-----/-------/--------/---------/--- ')
            df.to_excel("EstoqueMatriz.xlsx", index = False)


            with open("Balanço.csv", "a", newline="", encoding="utf-8") as arquivo:
                linhaa = dt.at[linha,"Descrição"]
                arquivo.write(f"{linhaa}, {peso_etiqueta_arrendodado}Kg \n ")
                arquivo.closed
            
        else:
            print("Linha não encontrada")

    def manual (self):
        
        df = pd.read_excel("EstoqueMatriz.xlsx")        
        
        codigo = int(self.codigo.get())
        peso = float(self.add_peso.get())
        
        self.limpa_manual()

        
        dt = df.loc[df["Código"] == codigo]
        print(dt)

        if not dt.empty:
            linha = dt.index[0]

            peso_anterior = float(df.at[linha, "estoque atual"])
            novo_peso = peso_anterior + peso

            print("Peso anterior: ",peso_anterior)
            print("Peso da etiqueta: ",peso)
            print("Novo peso na planilha: ",novo_peso)

            print('\n -----/-------/--------/---------/--- ')
            print('-----/-------/--------/---------/--- ')

            df.at[linha, "estoque atual"] = novo_peso
            df.to_excel("EstoqueMatriz.xlsx", index = False)


            with open("Balanço.csv", "a", newline="", encoding="utf-8") as arquivo:
                linhaa = dt.at[linha,"Descrição"]
                arquivo.write(f"{linhaa}, {peso}Kg \n ")
                arquivo.closed
            
        else:
            print("Linha não encontrada") 


    def manual_reducao (self):
        
        df = pd.read_excel("EstoqueMatriz.xlsx")        
        
        codigo = int(self.codigos.get())
        peso = float(self.add_pesos.get())
        self.limpa_manuals()

        
        dt = df.loc[df["Código"] == codigo]
        print(dt)

        if not dt.empty:
            linha = dt.index[0]

            peso_anterior = float(df.at[linha, "estoque atual"])
            novo_peso = peso_anterior - peso

            print("Peso anterior: ",peso_anterior)
            print("Peso da etiqueta: ",peso)
            print("Novo peso na planilha: ",novo_peso)

            print('\n -----/-------/--------/---------/--- ')
            print('-----/-------/--------/---------/--- ')

            df.at[linha, "estoque atual"] = novo_peso
            df.to_excel("EstoqueMatriz.xlsx", index = False)


            with open("Balanço.csv", "a", newline="", encoding="utf-8") as arquivo:
                linhaa = dt.at[linha,"Descrição"]
                arquivo.write(f"{linhaa}, {peso}Kg \n ")
                arquivo.closed
            
        else:
            print("Linha não encontrada")

    def zerar_planilha(self):
        df = pd.read_excel('EstoqueMatriz.xlsx')
        dt = df.loc[df["estoque atual"] > 0]

        if not dt.empty:
            df.loc[df["estoque atual"] > 0, "estoque atual"] = 0

            df.to_excel('EstoqueMatriz.xlsx', index = False)
            print("Estoque zerado")
            print('\n -----/-------/--------/---------/--- ')
            print('-----/-------/--------/---------/--- ')
        else:
            print('Não existe produto para se zerar')
            print('\n -----/-------/--------/---------/--- ')
            print('-----/-------/--------/---------/--- ')

janela = Tk()

class Tela(Arquivo):
    def __init__(self):
        self.tela = janela
        self.tela_()
        self.entrys()
        self.botao()

        janela.mainloop()

    def tela_(self):
        self.tela.title("Inventario de Estoque")
        self.tela.configure(background = "#255255266")
        self.tela.geometry("300x420")
        self.tela.resizable(True, True)
        self.tela.maxsize(width = 1350, height = 720)
        self.tela.minsize(width = 550, height = 250 )

    def entrys(self):
        self.add_product = Label(janela, text = "Código", bg = "#255255250", fg = "white", font = ("ariel", 8, "bold"))
        self.add_product.place(relx = 0.05, rely = 0.08)
        self.add_product = Entry(janela, bg = "#266200266", fg = "white")
        self.add_product.place(relx = 0.05, rely = 0.12, relheight = 0.06)

        self.reduzir_produto = Label(janela, text = "Código", bg = "#255255250", fg = "white", font = ("ariel", 8, "bold"))
        self.reduzir_produto.place(relx = 0.4, rely = 0.08)
        self.reduzir_produto = Entry(janela, bg = "#266200266", fg = "white")
        self.reduzir_produto.place(relx = 0.4, rely = 0.12, relheight = 0.06)

        self.codigo = Label(janela, text = "Código", bg = "#255255250", fg = "white", font = ("ariel", 8, "bold"))
        self.codigo.place(relx = 0.05, rely = 0.4)
        self.codigo = Entry(janela, bg = "#266200266", fg = "white")
        self.codigo.place(relx = 0.05, rely = 0.45, relheight = 0.06)

        self.add_peso = Label(janela, text = "Peso", bg = "#255255250", fg = "white", font = ("ariel", 8, "bold"))
        self.add_peso.place(relx = 0.05, rely = 0.55)
        self.add_peso = Entry(janela, bg = "#266200266", fg = "white")
        self.add_peso.place(relx = 0.05, rely = 0.60, relheight = 0.06)

        self.codigos = Label(janela, text = "Código", bg = "#255255250", fg = "white", font = ("ariel", 8, "bold"))
        self.codigos.place(relx = 0.5, rely = 0.4)
        self.codigos = Entry(janela, bg = "#266200266", fg = "white")
        self.codigos.place(relx = 0.5, rely = 0.45, relheight = 0.06)

        self.add_pesos = Label(janela, text = "Peso", bg = "#255255250", fg = "white", font = ("ariel", 8, "bold"))
        self.add_pesos.place(relx = 0.5, rely = 0.55)
        self.add_pesos = Entry(janela, bg = "#266200266", fg = "white")
        self.add_pesos.place(relx = 0.5, rely = 0.60, relheight = 0.06)


    def botao (self):
        self.somar = Button(janela, text = "Adicionar", command = self.somando)
        self.somar.place(relx = 0.05, rely = 0.20, relheight = 0.06)
        self.add_product.bind("<Return>", lambda event : self.somando())
        

        self.subtrair = Button(janela, text = "Remover", command = self.subtraindo)
        self.subtrair.place(relx = 0.4, rely = 0.20, relheight = 0.06)
        self.reduzir_produto.bind("<Return>", lambda event : self.subtraindo())


        self.subtrair = Button(janela, text ="Adicionar", command = self.manual)
        self.subtrair.place(relx = 0.30, rely = 0.53, relheight = 0.06)
        self.add_peso.bind("<Return>", lambda event : self.manual())

        self.subtrair = Button(janela, text ="Reduzir", command = self.manual_reducao)
        self.subtrair.place(relx = 0.75, rely = 0.53, relheight = 0.06)
        self.add_pesos.bind("<Return>", lambda event : self.manual_reducao())


        self.zerar = Button(janela, text ="Zerar Planilha", command = self.zerar_planilha)
        self.zerar.place(relx = 0.85, rely = 0.93, relheight = 0.06)
        self.zerar.bind("<Return>", lambda event : self.zerar_planilha())

Tela()