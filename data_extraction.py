import requests,json
import xlwings as xw
import pandas as pd

path = 'Motocycle.accident.variables.xlsx'

wb = xw.Book(path)
sheet = wb.sheets[0]

df = sheet['A4:v1000'].options(pd.DataFrame, index=False, header=True).value

states_code = pd.read_csv('states_code.csv')

url = 'https://crashviewer.nhtsa.dot.gov/CrashAPI'

for state in df ['State'].unique():
    total_accident = 0
    fatal = 0
    if state in states_code['State'].unique():
        state_id = states_code.loc[states_code['State']==state, 'Id'].values[0]

        api = '/analytics/GetInjurySeverityCounts?fromCaseYear=%s&toCaseYear=%s&state=%s&format=json' %(2009, 2019, state_id)
        r = requests.get(url+api)
        data_dic = r.json()
        for i in range(0, 9):
            print(data_dic['Results'][0][i])
            total_accident = total_accident + data_dic['Results'][0][i]['CrashCounts']
            fatal = fatal  + data_dic['Results'][0][i][
                'TotalFatalCounts']
        df.loc[df['State']==state, 'Motorcycle Accidents']=total_accident
        df.loc[df['State'] == state, 'Fatality #'] = fatal
    else:
        print(state)

df['Fatlity % '] = df['Fatality #']/df['Population']
df.to_csv('df.csv')
