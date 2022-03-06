import pandas as pd
import win32com.client as win32
import matplotlib.pyplot as plt

#importando os dados
vendas_df=pd.read_excel("Vendas.xlsx")
dez_df=pd.read_excel("Vendas - Dez.xlsx")

#adicionando o mes de dezembro ao nosso dataframe principal
vendas_df=vendas_df.append(dez_df)

#fazendo o histograma dos valores
vendas_df["Valor Final"].hist()
plt.title("valor final")
plt.savefig("hsitograma_final.png")
plt.show()

vendas_df["Valor Unitário"].hist()
plt.title("Valor Unitário")
plt.savefig("hsitograma_uni.png")
plt.show()

#determinando a quantidade de trasações por loja
trasacoes_por_loja=vendas_df["ID Loja"].value_counts.to_frame()

#criando um dataframe que contanhe a soma dos valores finais por loja
faturamento_loja=vendas_df[["ID Loja","Valor Final"]].groupby("ID Loja").sum()

#criando um dataframe que contanhe a soma dos valores finais por loja
quantidade_loja=vendas_df[["ID Loja","Quantidade"]].groupby("ID Loja").sum()

#verificando a media dos valor finais de acordo com a quantidade 
faturamento_medio=(faturamento_loja["Valor Final"]/quantidade_loja["Quantidade"]).to_frame()
faturamento_medio=faturamento_medio.rename(columns={0:"ticket medio"})

# local onde as imagens foram salvas
anexo1 = fr'.../hsitograma_final.png' 
anexo2 = fr'.../hsitograma_uni.png'

#enviando a analise feita por e-mail
outlook=win32.Dispatch("outlook.application") #conectando o python ao outlook do computador
mail=outlook.CreateItem(0) #criando um e-mail
mail.To="" #endereço  de email do destinatário
mail.Subject= " " #assunto do email
mail.Attachments.Add(anexo1)  #anexando a imagem 1
mail.Attachments.Add(anexo2)  #anexando a imagem 2
mail.HTMLBody= f'''<p>prezado,</p>

                    <h1>segue relatorio automatico dos dados</h1>

                    <p>faturament por loja</p> 
                    {faturamento_loja.to_html(formatters={"Valor Final":"R${:,.2f}".format})}

                    <p>quantidade vendida por loja,</p> 
                    {quantidade_loja.to_html()}

                    <p>valor medio das compras por loja</p> 
                    {faturamento_medio.to_html(formatters={"ticket medio":"R${:,.2f}".format})} 

                    <p>transações por loja </p> 
                    {trasacoes_por_loja.to_html()} 
                               ''' #corpo do email em linguagem html

mail.Send() #enviando o e-mail
