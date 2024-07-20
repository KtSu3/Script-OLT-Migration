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

#Função script

def model_ONT(PON, ONU, model, serial, name, desc, SLOT, vlan_calculada,):
  ont = (f'''

!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
!
interface gpon_olt-1/{SLOT}/{PON}
no onu {ONU}
!
interface gpon_olt-1/{SLOT}/{PON}
ONU {ONU} type {model} sn {serial}
!
interface gpon_ONU-1/{SLOT}/{PON}:{ONU}
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
pon-onu-mng gpon_onu-1/{SLOT}/{PON}:{ONU}
service DADOS-{vlan_calculada} gemport 1 vlan {vlan_calculada}
service VOIP-1600 gemport 2 vlan 1600
voip-ip ipv4 mode dhcp vlan-profile VOIP-1600 host 2
voip protocol sip
wan-ip 1 ipv4 mode pppoe username {name} password 123456789 vlan-profile WAN-{vlan_calculada} host 1
wan 1 mtu 1492 ethuni 1-4 ssid 1,5 service internet ipv6
wan-ip 1 ipv4 ping-resPONse enable traceroute-resPONse enable
sip-service pots_0/1 profile IP-PRIVATE userid userid password
security-mgmt 1 state enable mode forward ingress-type iphost 1 protocol web
security-mgmt 2 state enable mode forward ingress-type iphost 2 protocol web
!
!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    ''')
  with open("ONT.txt", "a") as archive:
       archive.write(ont)
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------#

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------#
       
# Iterar pelas linhas do DataFrame e imprimir linha por linha
for indice, linha in dados_excel.iterrows():
    ont = linha.values
    PON = int(ont[4])
    ONU = int(ont[5])
    desc = (ont[1]) 
    name = (ont[10]) 
    model = (ont[12])
    serial = (ont[6])
    SLOT = int(ont[3])
    usersip = (ont[13])
    password = (ont[14])
    vlan_calculada = calcular_vlan(SLOT, PON)
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------#

#Filtro ultilizando como referencia o Modelo dos equipamentos. 
    if model:
        match model:     
             case 'F6600PV9.0.12':
                  model_ONT(PON, ONU, model, serial, name, desc, SLOT, vlan_calculada,)
             case 'F670LV9.0':
                  model_ONT(PON, ONU, model, serial, name, desc, SLOT, vlan_calculada,)
             case 'SM164242-GHDUHR-T21':
                  model_ONT(PON, ONU, model, serial, name, desc, SLOT, vlan_calculada,)
             case 'R1.2.3':
                  model_ONT(PON, ONU, model, serial, name, desc, SLOT, vlan_calculada,)

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------#
# Contagem total de modelos do excel
count = dados_excel['Equipment Model'].value_counts().to_string()
with open ('Log_excel.txt',  'a') as archive:
     archive.write(count)



