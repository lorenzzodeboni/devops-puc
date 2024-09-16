#!/usr/bin/env python
# coding: utf-8

# ### Passo 1 - Importar Arquivos e Bibliotecas

# In[ ]:


import pandas as pd
import pathlib
import win32com.client as win32


# In[ ]:


emails = pd.read_excel(r"Bases de Dados\Emails.xlsx")
lojas = pd.read_csv("Bases de Dados\Lojas.csv", encoding="latin1", sep =";")
vendas = pd.read_excel("Bases de Dados\Vendas.xlsx")

display(emails)
display(lojas)
display(vendas)


# ### Passo 2 - Juntando as tabelas e definindo o dia indicador para trabalhar.

# In[ ]:


### Incluir nome da loja em vendas

vendas = vendas.merge(lojas, on= "ID Loja")

display(vendas)


# In[ ]:


### Criando uma tabela para cada uma das lojas:

dicionario_lojas = {}

for loja in lojas["Loja"]:
    
    dicionario_lojas[loja] = vendas.loc[vendas["Loja"] == loja, :]
    
display(dicionario_lojas["Rio Mar Recife"])



    
    


# In[ ]:


### Nesse caso, escolhi o dia indicador sendo o dia mais recente atualizando na planilha. 

### Para alterar a data desejada, basta inserir a data desejada no dia indicador na formatação 2019-mm-dd.

dia_indicador = vendas["Data"].max()

print("{}/{}".format(dia_indicador.day, dia_indicador.month))


# ### Passo 3 - Criando pastas para o backup de cada loja.

# In[ ]:


dia_indicador = vendas["Data"].max()



caminho_backup = pathlib.Path(r"Backup Arquivos Lojas")

arquivos_pasta_backup = caminho_backup.iterdir()

lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]

for loja in dicionario_lojas.keys():
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        
        nova_pasta.mkdir()
    
    
    ###criando arquivos backup
    nome_arquivo = "{}_{}_{}.xlsx".format(dia_indicador.month, dia_indicador.day, loja)

    local_arquivo = caminho_backup / loja / nome_arquivo
    
    dicionario_lojas[loja].to_excel(local_arquivo)
        
        
        
    
    


# ### Passo 4 - Calculando indicadores e automatizando os envios de e-mails.
# 

# In[ ]:


####### definindo as metas
meta_fat_dia = 1000
meta_fat_ano = 1650000
meta_prod_dia = 4
meta_prod_ano = 120
meta_ticket_dia = 500
meta_ticket_ano = 500


# In[ ]:




for loja in dicionario_lojas:

    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja["Data"] == dia_indicador, :]

    faturamento_ano = vendas_loja["Valor Final"].sum()
    faturamento_dia = vendas_loja_dia["Valor Final"].sum()


    #print(faturamento_ano)

    #print(faturamento_dia)

    ###diversidade de produtos
    qnt_ano = len(vendas_loja["Produto"].unique())
    qnt_dia = len(vendas_loja_dia["Produto"].unique())


    #print(qnt_ano)
    #print(qnt_dia)


    ###Ticket medio ano
    valor_venda_ano = vendas_loja.groupby("Código Venda").sum()

    ticket_medio_ano = valor_venda_ano["Valor Final"].mean()

    #print(ticket_medio_ano)

    ###Ticket medio dia
    valor_venda_dia = vendas_loja_dia.groupby("Código Venda").sum()

    ticket_medio_dia = valor_venda_dia["Valor Final"].mean()

    #print(ticket_medio_dia)



    ######### automatizando e-mail
    outlook = win32.Dispatch("outlook.application")

    nome = emails.loc[emails["Loja"] == loja, "Gerente"].values[0]
    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails["Loja"] == loja, "E-mail"].values[0]
    #mail.CC = 'email@gmail.com'
    #mail.BCC = 'email@gmail.com'
    mail.Subject = 'One page dia {}/{} - loja {}'.format(dia_indicador.day, dia_indicador.month, loja)
    #mail.Body = 'Texto do E-mail'





    cor1 = "green"
    cor2 = "red"

    if faturamento_dia >= meta_fat_dia:  
        cor_fat_dia = "green"
    else:
        cor_fat_dia = "red"

    if faturamento_ano >= meta_fat_ano:
        cor_fat_ano = "green"
    else:
        cor_fat_ano = "red"

    if qnt_dia >= meta_prod_dia:
        cor_qnt_dia = "green"
    else:
        cor_qnt_dia = "red"

    if qnt_ano >= meta_prod_ano:
        cor_qnt_ano = "green"
    else:
        cor_qnt_ano = "red"

    if ticket_medio_dia >= meta_ticket_dia:
        cor_ticket_mediodia = "green"
    else:
        cor_ticket_mediodia = "red"

    if ticket_medio_ano >= meta_ticket_ano:
        cor_ticket_medioano = "green"
    else:
        cor_ticket_medioano = "red"








    mail.HTMLBody = f"""   
    <p>Bom dia, {nome} </p>

    <p> O resultado de ontem <strong> ({dia_indicador.day}/{dia_indicador.month}) </strong> da loja <strong> {loja} </strong> foi: </p>

    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor dia</th>
        <th>Meta dia</th>
        <th>Cenário dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td> {faturamento_dia} </td>
        <td> {meta_fat_dia} </td>
        <th><font color ="{cor_fat_dia}">◙</th>
      </tr>
      <tr>
        <td>Diversidade de produtos</td>
        <td> {qnt_dia} </td>
        <td> {meta_prod_dia} </td>
        <th><font color ="{cor_qnt_dia}">◙</th>
      </tr>
      <tr>
        <td>Ticket medio</td>
        <td> {ticket_medio_dia} </td>
        <td> {meta_ticket_dia} </td>
        <th><font color ="{cor_ticket_mediodia}">◙</th>
      </tr>
    </table>
    <br>
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor dia</th>
        <th>Meta dia</th>
        <th>Cenário dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td> {faturamento_ano} </td>
        <td> {meta_fat_ano} </td>
        <th><font color ="{cor_fat_ano}">◙</th>
      </tr>
      <tr>
        <td>Diversidade de produtos</td>
        <td> {qnt_ano} </td>
        <td> {meta_prod_ano} </td>
        <th><font color ="{cor_qnt_ano}">◙</th>
      </tr>
      <tr>
        <td>Ticket medio</td>
        <td> {ticket_medio_ano} </td>
        <td> {meta_ticket_ano} </td>
        <th><font color ="{cor_ticket_medioano}">◙</th>
      </tr>
    </table>

    <p> Segue em anexo a planilha com todos os dados para análise. </p>

    <p> Att, </p>

    <p> Lorenzzo Deboni </p>






    """

    # Anexos (pode colocar quantos quiser):
    attachment  = pathlib.Path.cwd() / caminho_backup / loja / "{}_{}_{}.xlsx".format(dia_indicador.month, dia_indicador.day,loja)
    mail.Attachments.Add(str(attachment))

    mail.Send()

    
    print("E-mail da loja {} enviado com sucesso".format(loja))



    
    


# ### Passo 7 - Criar ranking para diretoria

# In[ ]:


### Ranking para o ano

faturamento_lojas = vendas.groupby("Loja").sum()

faturamento_lojas = faturamento_lojas.reset_index()

faturamento_lojas =  faturamento_lojas[["Loja", "Valor Final"]]

faturamento_lojas_ano = faturamento_lojas.sort_values(by = "Valor Final", ascending = False)

display(faturamento_lojas_ano)

### Passando a tqabela rankind anual para excel.
nome_arquivo = "{}_{}_Ranking Anual.xlsx".format(dia_indicador.month, dia_indicador.day)

faturamento_lojas_ano.to_excel(r"Backup Arquivos Lojas\{}".format(nome_arquivo))

### Ranking para o dia

vendas_dia = vendas.loc[vendas["Data"] == dia_indicador, :]

vendas_dia = vendas_dia.groupby("Loja").sum()

vendas_dia = vendas_dia.reset_index()

faturamento_lojas_dia = vendas_dia[["Loja", "Valor Final"]]

faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by = "Valor Final", ascending = False)

display(faturamento_lojas_dia)


#### Passando a tabela ranking dia para o excel.

nome_arquivo = "{}_{}_Ranking Dia.xlsx".format(dia_indicador.month, dia_indicador.day)

faturamento_lojas_dia.to_excel(r"Backup Arquivos Lojas\{}".format(nome_arquivo))


# ### Passo 8 - Enviar e-mail para diretoria

# In[ ]:


######### automatizando e-mail
outlook = win32.Dispatch("outlook.application")


mail = outlook.CreateItem(0)
mail.To = emails.loc[emails["Loja"] == "Diretoria", "E-mail"].values[0]
#mail.CC = 'email@gmail.com'
#mail.BCC = 'email@gmail.com'
mail.Subject = 'Ranking Dia {}/{}'.format(dia_indicador.day, dia_indicador.month)
mail.Body = '''

Prezados, bom dia!

Segue em anexo os ranking do ano e do dia de todas as lojas.'''


# Anexos (pode colocar quantos quiser):
attachment  = pathlib.Path.cwd() / caminho_backup / "{}_{}_Ranking Anual.xlsx".format(dia_indicador.month, dia_indicador.day)
mail.Attachments.Add(str(attachment))
attachment  = pathlib.Path.cwd() / caminho_backup / "{}_{}_Ranking Dia.xlsx".format(dia_indicador.month, dia_indicador.day)
mail.Attachments.Add(str(attachment))
mail.Send()

