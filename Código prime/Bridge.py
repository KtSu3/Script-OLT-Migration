import pandas as pd
counter = 0

#Extração de dados do Excel
dados_excel = pd.read_excel(r'C:\Users\kauam\py\script-migracao-em-massa-zte\Sheets\1204.xlsx')
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
def model_bridge(PON, ONU, model, serial, name, desc, SLOT, vlan_calculada):
     bridge = (f'''
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
gemport 1 name DADOS-{vlan_calculada} tcont 1
!
interface vport-1/{SLOT}/{PON}.{ONU}:1
service-port 1 user-vlan {vlan_calculada} vlan {vlan_calculada} new-cos 0 svlan 1100
!
pon-onu-mng gpon_onu-1/{SLOT}/{PON}:{ONU}
service DADOS-{vlan_calculada} gemport 1 vlan {vlan_calculada}
vlan port eth_0/1 mode tag vlan {vlan_calculada}
vlan port eth_0/2 mode tag vlan {vlan_calculada}
vlan port eth_0/3 mode tag vlan {vlan_calculada}
vlan port eth_0/4 mode tag vlan {vlan_calculada}
!
!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
''')
     with open("Bridge.txt", "a") as archive:
       archive.write(bridge)
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------#
       
# Iterar pelas linhas do DataFrame e imprimir linha por linha
for indice, linha in dados_excel.iterrows():
    bridge = linha.values
    PON = int(bridge[4])
    ONU = int(bridge[5])
    desc = (bridge[1]) 
    name = (bridge[10]) 
    model = (bridge[12])
    serial = (bridge[6])
    SLOT = int(bridge[3])
    usersip = (bridge[13])
    password = (bridge[14])
    vlan_calculada = calcular_vlan(SLOT, PON)
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------#

#Filtro ultilizando como referencia o Modelo dos equipamentos. 
    if model:
        match model:
             case 'AN5506_01A1':
                  model_bridge(PON, ONU, model, serial, name, desc, SLOT, vlan_calculada)
             case 'F601V7.0':
                  model_bridge(PON, ONU, model, serial, name, desc, SLOT, vlan_calculada)
             case 'TX-6610':
                  model_bridge(PON, ONU, model, serial, name, desc, SLOT, vlan_calculada)
             case 'XZ000-G3':
                  model_bridge(PON, ONU, model, serial, name, desc, SLOT, vlan_calculada)
             case 'DM986-100':
                  model_bridge(PON, ONU, model, serial, name, desc, SLOT, vlan_calculada)

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------#
