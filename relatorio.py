import pandas as pd
import win32com.client as win32
import matplotlib.pyplot as plt

#importando os dados
vendas_df=pd.read_excel("Vendas.xlsx")
dez_df=pd.read_excel("Vendas - Dez.xlsx")

#adicionando o mes de dezembro ao nosso datafreime principla
vendas_df=vendas_df.append(dez_df)

#outra forma de terderminar a corelação entre as colunas é utilizando o pandas
vendas_df["Valor Final"].hist()
plt.title("valor final")
plt.savefig("hsitograma_final.png")
plt.show()

vendas_df["Valor Unitário"].hist()
plt.title("Valor Unitário")
plt.savefig("hsitograma_uni.png")
plt.show()

#determinando a quantidade de trasações por loja
trasacoes_por_loja=vendas_df["ID Loja"].value_counts()

#criando um dataframe que contanhe a soma dos valores finais por loja
faturamento_loja=vendas_df[["ID Loja","Valor Final"]].groupby("ID Loja").sum()

#criando um dataframe que contanhe a soma dos valores finais por loja
quantidade_loja=vendas_df[["ID Loja","Quantidade"]].groupby("ID Loja").sum()

#verificando a media dos valor finais de acordo com a quantidade 
ticket_medio=(faturamento_loja["Valor Final"]/quantidade_loja["Quantidade"]).to_frame()
ticket_medio=ticket_medio.rename(columns={0:"ticket medio"})

# enviando a analise feita por e-mail

anexo1 = fr'C:/Users/Csleo/Desktop/IFUSP/python/python/pandas/hsitograma_final.png'
anexo2 = fr'C:/Users/Csleo/Desktop/IFUSP/python/python/pandas/hsitograma_uni.png'

outlook=win32.Dispatch("outlook.application") #conectando o python ao outlook do computador
mail=outlook.CreateItem(0) #criando um e-mail
mail.To="l.c.pegasus12@gmail.com" #para quem vai o email
mail.Subject= " teste do codigo " #assunto do email
mail.Attachments.Add(anexo1)
mail.Attachments.Add(anexo2)
mail.HTMLBody= f'''<p>prezado,</p>

                    <h1>segue relatorio automatico dos dados</h1>

                    <p>telabe do faturament</p> 
                    {faturamento_loja.to_html(formatters={"Valor Final":"R${:,.2f}".format})}

                    <p>tabela da quantidade vendida por loja,</p> 
                    {quantidade_loja.to_html()}

                    <p>tabela do ticket medio </p> 
                    {ticket_medio.to_html(formatters={"ticket medio":"R${:,.2f}".format})} 

                    <p>transações por loja </p> 
                    {trasacoes_por_loja.to_html()} 
                               ''' #corpo do email em linguagem html

mail.Send() #enviando o e-mail
