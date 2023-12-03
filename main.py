import pandas as pd
import datetime

datum = datetime.datetime.now().strftime("%m-%Y")

csv = pd.read_csv('overview.csv', delimiter=';')

result = pd.DataFrame(columns=['Firma', 'Essential', 'Business', 'Enterprise', 'Enterprise-Dial-in-pack', 'Attendant', 'Alert', 'Connect', 'Dial in per use', 'Room', 'Business-Custom', 'Enterprise-Custom', 'Attendant-Custom', 'Voice-Business-Custom', 'Voice-Enterprise-Custom', 'Voice-Attendant-Custom', 'Alert-Custom', 'Room-Custom'])

for i, row in csv.iterrows():
    customer_name = row['customerName']
    volume = row['volume']
    lizenz_typ = 'NA'

    if 'Essential' in row['rainbowServiceId']:
        lizenz_typ = 'Essential'

    elif 'Business' in row['rainbowServiceId']:
        lizenz_typ = 'Business'

    elif 'Voice-Enterprise-Custom' in row['rainbowServiceId']:
        lizenz_typ = 'Voice-Enterprise-Custom'

    elif 'Enterprise-Custom' in row['rainbowServiceId']:
        lizenz_typ = 'Enterprise-Custom'

    elif 'Enterprise-Dial-in-pack' in row['rainbowServiceId']:
        lizenz_typ = 'Enterprise-Dial-in-pack'

    elif 'Enterprise' in row['rainbowServiceId']:
        lizenz_typ = 'Enterprise'

    elif 'Alert-Custom' in row['rainbowServiceId']:
        lizenz_typ = 'Alert-Custom'

    elif 'Room-Custom' in row['rainbowServiceId']:
        lizenz_typ = 'Room-Custom'

    elif 'Alert' in row['rainbowServiceId']:
        lizenz_typ = 'Alert'

    elif 'Connect' in row['rainbowServiceId']:
        lizenz_typ = 'Connect'

    elif 'Dial-in-per-use' in row['rainbowServiceId']:
        lizenz_typ = 'Dial-in-per-use'

    elif 'Room' in row['rainbowServiceId']:
        lizenz_typ = 'Room'

    elif 'Voice-Attendant-Custom' in row['rainbowServiceId']:
        lizenz_typ = 'Voice-Attendant-Custom'

    elif 'Attendant-Custom' in row['rainbowServiceId']:
        lizenz_typ = 'Attendant-Custom'

    elif 'Attendant' in row['rainbowServiceId']:
        lizenz_typ = 'Attendant'

    if customer_name in result["Firma"]:
        result.loc[result['Firma'] == customer_name, f'{lizenz_typ}'] +=volume
    else:
        new_row = {'Firma': customer_name}
        new_row.update({f'{lt}': volume if lt == lizenz_typ else 0 for lt in result.columns[1:]})
        result = result._append(new_row, ignore_index=True)

result = result.groupby('Firma').sum()
result.to_csv(f'Lizenzen_{datum}.csv', sep=';', index=True, encoding='utf-8')
