import pandas as pd
counter = 0

#Extração de dados do Excel
dados_excel = pd.read_excel(r'C:\Users\kauam\py\script-migracao-em-massa-zte\Sheets\1205.xlsx')
coluna_para_verificar = 'PON Number'
dados_excel = dados_excel.dropna(subset=[coluna_para_verificar])
novo_arquivo_excel = 'novo_arquivo.xlsx'
dados_excel.to_excel(novo_arquivo_excel, index=False)
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------#

#Função para Calcular a Vlan
def calcular_vlan(SLOT,PON):
    resultado = int((SLOT * 100) + PON + 20)
    return resultado
#-------------------------------------------------------------------------------------------q---------------------------------------------------------------------------#


#Funções para retornar o Script, cada modelo trás sua função, Bridge, ONT, Bridge_Voip
def model_voip(PON, ONU, model, serial, name, desc, SLOT, vlan_calculada, usersip, password):
  voip = (f'''
!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
!
interface gpon_olt-1/{SLOT}/{PON}
no onu {ONU}
!
interface gpon_olt-1/{SLOT}/{PON}
onu {ONU} type {model} sn {serial}
!
interface gpon_onu-1/{SLOT}/{PON}:{ONU}
name {name}
description {desc}
tcont 1 profile 1G
tcont 2 profile VOIP-256K
gemport 1 name DADOS-{vlan_calculada} tcont 1
gemport 2 name VOIP-1600 tcont 2
!
interface vport-1/{SLOT}/{PON}.{ONU}:1
service-port 1 user-vlan {vlan_calculada} vlan {vlan_calculada} new-cos 0 svlan 1100 new-cos 0
!
interface vport-1/{SLOT}/{PON}.{ONU}:2
service-port 2 user-vlan 1600 vlan 1600 new-cos 0
!
pon-onu-mng gpon_onu-1/{SLOT}/{PON}:1
service DADOS-{vlan_calculada} gemport 1 vlan {vlan_calculada}
service VOIP-1600 gemport 2 vlan 1600
voip-ip ipv4 mode dhcp vlan-profile VOIP-1600 host 1
voip protocol sip
vlan port eth_0/1 mode tag vlan {vlan_calculada}
vlan port eth_0/2 mode tag vlan {vlan_calculada}
vlan port eth_0/3 mode tag vlan {vlan_calculada}
vlan port eth_0/4 mode tag vlan {vlan_calculada}
sip-service pots_0/1 profile IP-PRIVATE userid 
sip-service pots_0/2 profile  IP-PRIVATE userid 
!
!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
''')
  with open("Voip.txt", "a") as archive:
       archive.write(voip)
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------#
       
# Iterar pelas linhas do DataFrame e imprimir linha por linha
for indice, linha in dados_excel.iterrows():
    if linha.values == (linha.values.isnull().sum()):
     voip = linha.values
     PON = int(voip[4])
     ONU = int(voip[5])
     desc = (voip[1]) 
     name = (voip[10]) 
     model = (voip[12])
     serial = (voip[6])
     SLOT = int(voip[3])
     usersip = int(voip[13])
     password = str(voip[14])
     vlan_calculada = calcular_vlan(SLOT, PON)
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Filtro ultilizando como referencia o Modelo dos equipamentos. 
    if model:
        match model:
             case 'AN5506_02B':
                model_voip(PON, ONU, model, serial, name, desc, SLOT, vlan_calculada,usersip,password)
             case 'AN5506_04B2':
                model_voip(PON, ONU, model, serial, name, desc, SLOT, vlan_calculada,usersip,password)
             case 'AN5506_04F1':
                model_voip(PON, ONU, model, serial, name, desc, SLOT, vlan_calculada,usersip,password)
             case 'F612V9.2':
                model_voip(PON, ONU, model, serial, name, desc, SLOT, vlan_calculada,usersip,password)
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------





