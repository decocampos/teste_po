from Authentication import GoogleAuthentication
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from notification import NotificationHandler 
import requests

#===============================================================

'''CONFIGURAÇÃO PLANILHA'''
#Informações para o acesso da planilha
#Com 'SAMPLE_SPREADSHEET_ID' sendo o ID da planilha utilizada
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SAMPLE_SPREADSHEET_ID = "104nykhGL5aElGJ0JYx-Mda67Hd0rgOoKTxF-XghF_S8"

#Liberando acesso à planilha para modificações
google_auth=GoogleAuthentication(SCOPES)
creds= google_auth.authenticate()



#===============================================================

'''FUNÇÕES ÚTEIS'''

def update_spreadsheet(sheets, i, j, value):
    if isinstance(value, dict):
        value_str = '; '.join([f"{key}: {value:.2f}" for key, value in value.items()])
        value = value_str
    
    sheets.values().update(
        spreadsheetId=SAMPLE_SPREADSHEET_ID,
        range=f"Projetos!{j}{i}",
        valueInputOption="USER_ENTERED",
        body={"values": [[value]]}
    ).execute()


                    donator[name_donator] = value_donator
                j+=1
            #Colocando as informações na planilha
            update_spreadsheet(sheets, i,'I',donator )
            update_spreadsheet(sheets, i,'J', sum_value)
    else:
        print("Erro ao fazer a requisição:", response.status_code)
    i+=1
notification.notify_concluded("CÓDIGO FINALIZADO","Seu código foi finalizado com sucesso! Agora você pode fechar sua IDE, e utilizar a planilha com os dados todos funcionais")

