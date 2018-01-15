#%%
from openpyxl import load_workbook, Workbook
import datetime

######################
# Актуализация списка сотрудников

# ссылка
url = 'test.xlsx'

# файл
file = load_workbook(filename = url)

# листы файла
fileSheets = file.sheetnames

# Градация баллов
pointsSheet = file['ТП']

pointsDict = {}

for i in range(2, len(pointsSheet['A']) + 1):
    tp = pointsSheet['A%d' % i].value
    p = pointsSheet['C%d' % i].value
    pointsDict.setdefault(tp, [p, i])
    
print(pointsDict)
print()

# Сотрудники (с 0 баллами по умолчанию)
workersSheet = file['сотрудники']

workersDict = {}


for i in range(2, len(workersSheet['C']) + 1):
    dvo = workersSheet['C%d' % i].value
    fio = workersSheet['F%d' % i].value
    workersDict.setdefault(dvo, [fio, 0])
    for tp in pointsDict:
        workersDict[dvo].append(0)
        
print(workersDict)
    
    
#%%  
# Операции (самое долгое)
operatinos = load_workbook(filename = 'fortest.xlsx')
operationsSheet = operatinos['исх']

for i in range(2, len(operationsSheet['AA']) + 1):
    dvo = operationsSheet['AA%d' % i].value
    tp = operationsSheet['H%d' % i].value    
    if dvo in workersDict and tp in pointsDict:
        workersDict[dvo][1] += pointsDict[tp][0]
        workersDict[dvo][pointsDict[tp][1]] += pointsDict[tp][0]


#%%        
'''
получается, что словарь workers становится консалидированным
списком, где показывается скольк у каждого сотрудника баллов всего
и сколько за каждый ТП, 
теперь можно просто этот словарь записать в новый файл
'''

# осталось словарь перенести в новый файл
boiNew = Workbook()
boiNewSheet = boiNew.active

for dvo in workersDict:
    boiNewSheet.append(workersDict[dvo])

now = datetime.datetime.today().strftime('%d.%m.%Y %H.%M')
        
boiNew.save(now +'.xlsx')
            
        







  


        

