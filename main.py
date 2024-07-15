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

#===============================================================
'''CÓDGIGO PRINCIPAL'''
notification = NotificationHandler()
notification.notify_concluded("CÓDIGO INICIADO","Seu código foi iniciado, agora, você pode acompanhar a planilha sendo modificada. Assim que for concluído, você receberá outra notificação dessa, enquanto isso deixa sua máquina ligada")

i=2 #Contador para escrever a posição correta na planilha
for pronac in cell_value:
    donator=''
    sum_value=0
    #Cada PRONAC que tem na planilha é consultado e retirado as informação
    url = f"http://api.salic.cultura.gov.br/v1/projetos/{pronac[0]}?format=json"
    response = requests.get(url)

    if response.status_code == 200:
        data = response.json()
        #Situação caso, ele não tenha recebido valor dos patrocinadores
        if len(data['_embedded']['captacoes'])==0:
            update_spreadsheet(sheets, i,'I', "Não possui patrocinadores")
        #Caso ele possua patrocinadores
        else:
            j=0 # Contador para caso exista mais de um patrocinador
            donator={}
            for donator_info in data['_embedded']['captacoes']:
                name_donator= str(data['_embedded']['captacoes'][j]['nome_doador']) 
                value_donator=float(data['_embedded']['captacoes'][j]['valor'])
                sum_value+=value_donator
                if name_donator in donator:
                    donator[name_donator] += value_donator
                else:
                    donator[name_donator] = value_donator
                j+=1
            #Colocando as informações na planilha
            update_spreadsheet(sheets, i,'I',donator )
            update_spreadsheet(sheets, i,'J', sum_value)
    else:
        print("Erro ao fazer a requisição:", response.status_code)
    i+=1
notification.notify_concluded("CÓDIGO FINALIZADO","Seu código foi finalizado com sucesso! Agora você pode fechar sua IDE, e utilizar a planilha com os dados todos funcionais")


