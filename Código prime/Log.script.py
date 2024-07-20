import pandas as pd

models = {
    'ont_f660p': 'F6600PV9.0.12',
    'ont_f670fodase': 'F670LV9.0',
    'ont_sumec': 'SM164242-GHDUHR-T21',
    'ont_sumec2': 'R1.2.3',
 }

teste = 'F6600PV9.0.12'

with open(r'C:\Users\kauam\py\script-migracao-em-massa-zte\venv\Código prime\ONT.txt', 'r') as archive:
    key = archive.read().strip()


if key in models:
   valor = teste
    

print(valor)