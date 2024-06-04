import pandas as pd
from pathlib import Path
import os
import requests
import time
import json


erros = 0
enviados = 0


##Configuração do endpoint responsavel pelo envio da mensagem no whatsapp
endpoint = "https://6bfbce4dd9d35.tunnel.zaplink.net/mensagem/envio"
headers = {
    "Authorization": "Bearer Token1", 
    "Content-Type": "application/json"
}

#Leitura do arquivo contendo os numeros e nomes para envio da msg
file_disparo = os.path.join(Path.home(), 'Downloads\DDD_Faltantes.xlsx')
df = pd.read_excel(file_disparo)

def enviar_mensagem(numero, mensagem):
    body = json.dumps({ 
        "number": numero,
        "body": mensagem
    })
    resultado = requests.post(endpoint, headers=headers, data=body)
    return resultado

#Passando por cada linha do excel (cada linha é um número de celular)
for index, row in df.iterrows():
    if index != 0:  
        time.sleep(20) 
    
    mensagem_personalizada = f"Olá, {row['Nome Pessoa']}\nSeu DDD é {ddd}!"
    
    telefone = str(row['Telefone'])
    ddd = telefone[:3]
    
    #Função responsável pela aplicação do DRY
    resultado = enviar_mensagem(telefone, mensagem_personalizada)


    #Controle dos envios da mensagem OK/Erro
    if 'OK' in resultado:
        print(f"ENVIADO - {row['Nome Pessoa']}: {resultado} - {row['Telefone']}")
        enviados += 1
        df.at[index, 'STATUS'] = 'ENVIO'
    else:
        print(f"ERRO - {row['Nome Pessoa']}: {resultado} - {row['Telefone']}")
        erros += 1
        df.at[index, 'STATUS'] = 'ERRO'
    

#Salvamento para visualização dos status dos envios
print(f'\n\nENVIOS: {enviados}\nERROS: {erros}')
df.to_excel('Status dos Envios.xlsx',index=False)
