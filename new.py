from importlib.resources import path
from new_func import *
import pandas as py
from openpyxl import load_workbook

day_range = input('day range(ex: 1-1 || 1-23): ')
month = input('number of the current month: ').strip()
year = input('which year are you currently in:').strip()

days = [int(day_range.split('-')[0].strip()),int(day_range.split('-')[1].strip())]
day_upper = max(days)
day_lower = min(days)
folder_path = r'C:\Users\teamf\Desktop'
dfmain_1 = pd.read_excel(r'C:\Users\teamf\Desktop\a.xlsx', sheet_name= 'VIDAPIRESVEOVALAMA')
dfmain_2 = pd.read_excel(r'C:\Users\teamf\Desktop\a.xlsx', sheet_name= 'MONTAJ&PAKETLEME')

for day in range(day_lower,day_upper+1):
    print('Gün: '+str(day))
    try:
        varied_path_1 = f'\{day}.{month}.{year} FABRİKA 1.xlsx'
        path_factory_1 = folder_path+varied_path_1
        gday = 'G'+str(day)
        dfmain_1 = factory_1(gday,rf'{path_factory_1}',dfmain_1)
    except (FileNotFoundError):
        print('Bu dosya bulunamadı bir sonraki geceye geçiliyor...')

    try:
        varied_path_2 = f'\{day}.{month}.{year} FABRİKA 2.xlsx'
        path_factory_2 = folder_path+varied_path_2
        gday = 'G'+str(day)
        dfmain_1 = factory_2(gday,rf'{path_factory_2}',dfmain_1)
    except (FileNotFoundError):
        print('Bu dosya bulunamadı bir sonraki güne geçiliyor...')

    try:
        varied_path_3 = '\AĞUSTOS MONTAJ 2022 (1).xlsx'
        path_montaj = folder_path + varied_path_3
        dfmain_2 = montaj(str(day),month,year,rf'{path_montaj}',dfmain_2)
    except (ValueError):
        print('Montajda öyle bir gün bulunamadı.')

FilePath = r'C:\Users\teamf\Desktop\b.xlsx'
with pd.ExcelWriter(FilePath, engine = 'openpyxl', mode = "a") as writer:
    dfmain_1.to_excel(writer, sheet_name = 'VIDAPIRESVEOVALAMA', index = False)

FilePath = r'C:\Users\teamf\Desktop\b.xlsx'
with pd.ExcelWriter(FilePath, engine = 'openpyxl', mode = "a") as writer:
    dfmain_2.to_excel(writer, sheet_name = 'MONTAJ&PAKETLEME', index = False)

print("bitti")
input()
