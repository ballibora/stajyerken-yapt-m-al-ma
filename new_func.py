import pandas as pd
from openpyxl import load_workbook

def factory_2(day,path_1,dfmain):    

    df_1 = pd.read_excel(path_1, sheet_name= 'PRES FABRİKA 2  - GÜNDÜZ', skiprows=151, skipfooter=33, usecols='D:L', index_col= 0)
    df_2 = pd.read_excel(path_1, sheet_name= 'PRES FABRİKA 2  - GÜNDÜZ', skiprows=162, skipfooter=22, usecols='D:L', index_col= 0)
    df_3 = pd.read_excel(path_1, sheet_name= 'PRES FABRİKA 2  - GÜNDÜZ', skiprows=173, skipfooter=11, usecols='D:L', index_col= 0)
    df_4 = pd.read_excel(path_1, sheet_name= 'PRES FABRİKA 2  - GÜNDÜZ', skiprows=184, skipfooter=0, usecols='D:L', index_col= 0)
    df_5 = pd.read_excel(path_1, sheet_name= 'PRES FABRİKA 2  - GECE', skiprows=150, skipfooter=33, usecols='C:K', index_col= 0)
    df_6 = pd.read_excel(path_1, sheet_name= 'PRES FABRİKA 2  - GECE', skiprows=161, skipfooter=22, usecols='C:K', index_col= 0)
    df_7 = pd.read_excel(path_1, sheet_name= 'PRES FABRİKA 2  - GECE', skiprows=172, skipfooter=11, usecols='C:K', index_col= 0)
    df_8 = pd.read_excel(path_1, sheet_name= 'PRES FABRİKA 2  - GECE', skiprows=183, skipfooter=0, usecols='C:K', index_col= 0)
    df_9 = pd.read_excel(path_1, sheet_name= 'OVALAMA FABRİKA 2  - GÜNDÜZ', skiprows=118, skipfooter=21, usecols='A:J', index_col= 0)
    df_10 = pd.read_excel(path_1, sheet_name= 'OVALAMA FABRİKA 2  - GÜNDÜZ', skiprows=129, skipfooter=10, usecols='A:J', index_col= 0)
    df_11 = pd.read_excel(path_1, sheet_name= 'OVALAMA FABRİKA 2  - GÜNDÜZ', skiprows=118, skipfooter=21, usecols='O:X', index_col= 0)
    df_12 = pd.read_excel(path_1, sheet_name= 'OVALAMA FABRİKA 2  - GÜNDÜZ', skiprows=129, skipfooter=10, usecols='O:W', index_col= 0)
    df_13 = pd.read_excel(path_1, sheet_name= 'OVALAMA FABRİKA 2  - GECEE', skiprows=119, skipfooter=18, usecols='O:W', index_col= 0)
    df_14 = pd.read_excel(path_1, sheet_name= 'OVALAMA FABRİKA 2  - GECEE', skiprows=130, skipfooter=7, usecols='O:W', index_col= 0)
    df_15 = pd.read_excel(path_1, sheet_name= 'OVALAMA FABRİKA 2  - GECEE', skiprows=119, skipfooter=18, usecols='A:J', index_col= 0)
    df_16 = pd.read_excel(path_1, sheet_name= 'OVALAMA FABRİKA 2  - GECEE', skiprows=130, skipfooter=7, usecols='A:J', index_col= 0)
    df_17 = pd.read_excel(path_1, sheet_name= 'PRES FABRİKA 2  - GÜNDÜZ', skiprows=141, skipfooter= 45, usecols='H:I', index_col= 0)
    df_18 = pd.read_excel(path_1, sheet_name= 'PRES FABRİKA 2  - GECE', skiprows=140, skipfooter=44, usecols='F:G', index_col= 0)
    df_19 = pd.read_excel(path_1, sheet_name= 'OVALAMA FABRİKA 2  - GÜNDÜZ', skiprows=120, skipfooter=21, usecols='L:M', index_col= 0)
    df_20 = pd.read_excel(path_1, sheet_name= 'OVALAMA FABRİKA 2  - GECEE', skiprows=120, skipfooter=18, usecols='L:M', index_col= 0)


    for table in range(1,17):
        for j in range(0,len(locals()[f'df_{table}'].loc['OPERATÖR'])):
            if pd.isna(locals()[f'df_{table}'].loc['OPERATÖR'][j]) == 0:
                if 'Unnamed:' not in locals()[f'df_{table}'].loc['OPERATÖR'][j] and 'TOPLAM' not in locals()[f'df_{table}'].loc['OPERATÖR'][j]:
                    dfmain.loc[dfmain.PERSONEL == locals()[f'df_{table}'].loc['OPERATÖR'][j].strip(), day] = locals()[f'df_{table}'].loc['ÜRETİM YÜZDESİ'][j]
        if pd.isnull(locals()[f'df_{table}'].index[0]) == 0:
            if '&' in locals()[f'df_{table}'].index[0]:
                names = locals()[f'df_{table}'].index[0].split('&')
                for name in names:
                    dfmain.loc[dfmain.PERSONEL == name.strip(), day] = locals()[f'df_{table}'].loc['ÜRETİM YÜZDESİ'][-1]
            else:
                dfmain.loc[dfmain.PERSONEL == locals()[f'df_{table}'].index[0].strip(), day] = locals()[f'df_{table}'].loc['ÜRETİM YÜZDESİ'][-1]
        if table in range(1,5):    
            if pd.isnull(locals()[f'df_{table}'].index.name) == 0:
                dfmain.loc[dfmain.PERSONEL == locals()[f'df_{table}'].index.name.strip(), day] = df_17.loc['ÜRETİM YÜZDESİ'][0]
        elif table in range(5,9):
            if pd.isnull(locals()[f'df_{table}'].index.name) == 0:
                dfmain.loc[dfmain.PERSONEL == locals()[f'df_{table}'].index.name.strip(), day] = df_18.loc['ÜRETİM YÜZDESİ'][0]
        elif table in range(9,13):
            if pd.isnull(locals()[f'df_{table}'].index.name) == 0:
                dfmain.loc[dfmain.PERSONEL == locals()[f'df_{table}'].index.name.strip(), day] = df_19.loc['ÜRETİM YÜZDESİ'][0]
        elif table in range(13,17):
            if pd.isnull(locals()[f'df_{table}'].index.name) == 0:
                dfmain.loc[dfmain.PERSONEL == locals()[f'df_{table}'].index.name.strip(), day] = df_20.loc['ÜRETİM YÜZDESİ'][0]
    return dfmain
def factory_1(day,path_1,dfmain):    

    #işçilerin performanslarını almak için çektiğim tablolar
    df_1 = pd.read_excel(path_1, sheet_name= 'PRES  - GÜNDÜZ (3)', skiprows=132, skipfooter=20, usecols='B:H', index_col= 0)
    df_2 = pd.read_excel(path_1, sheet_name= 'PRES  - GÜNDÜZ (3)', skiprows=142, skipfooter=10, usecols='B:H', index_col= 0)
    df_3 = pd.read_excel(path_1, sheet_name= 'PRES  - GÜNDÜZ (3)', skiprows=152, usecols='B:H', index_col= 0)
    df_4 = pd.read_excel(path_1, sheet_name= 'PRES  - GÜNDÜZ (3)', skiprows=132, skipfooter=20, usecols='M:U', index_col= 0)
    df_5 = pd.read_excel(path_1, sheet_name= 'PRES  - GÜNDÜZ (3)', skiprows=142, skipfooter=10, usecols='M:U', index_col= 0)
    df_6 = pd.read_excel(path_1, sheet_name= 'PRES  - GÜNDÜZ (3)', skiprows=152, usecols='M:U', index_col= 0)
    df_7 = pd.read_excel(path_1, sheet_name= 'PRES  - GECE (3)', skiprows=137, skipfooter=22, usecols='C:I', index_col= 0)
    df_8 = pd.read_excel(path_1, sheet_name= 'PRES  - GECE (3)', skiprows=147, skipfooter=12, usecols='C:I', index_col= 0)
    df_9 = pd.read_excel(path_1, sheet_name= 'PRES  - GECE (3)', skiprows=157, skipfooter=2, usecols='C:I', index_col= 0)
    df_10 = pd.read_excel(path_1, sheet_name= 'PRES  - GECE (3)', skiprows=137, skipfooter=22, usecols='N:T', index_col= 0)
    df_11 = pd.read_excel(path_1, sheet_name= 'PRES  - GECE (3)', skiprows=147, skipfooter=12, usecols='N:T', index_col= 0)
    df_12 = pd.read_excel(path_1, sheet_name= 'PRES  - GECE (3)', skiprows=157, skipfooter=2, usecols='N:T', index_col= 0)
    df_13 = pd.read_excel(path_1, sheet_name= 'OVALAMA - GÜNDÜZ (2)', skiprows=104, skipfooter=41, usecols='B:J', index_col= 0)
    df_14 = pd.read_excel(path_1, sheet_name= 'OVALAMA - GÜNDÜZ (2)', skiprows=114, skipfooter=31, usecols='B:K', index_col= 0)
    df_15 = pd.read_excel(path_1, sheet_name= 'OVALAMA - GÜNDÜZ (2)', skiprows=124, skipfooter=21, usecols='B:K', index_col= 0)
    df_16 = pd.read_excel(path_1, sheet_name= 'OVALAMA - GÜNDÜZ (2)', skiprows=104, skipfooter=41, usecols='P:W', index_col= 0)
    df_17 = pd.read_excel(path_1, sheet_name= 'OVALAMA - GÜNDÜZ (2)', skiprows=114, skipfooter=31, usecols='P:Y', index_col= 0)
    df_18 = pd.read_excel(path_1, sheet_name= 'OVALAMA - GÜNDÜZ (2)', skiprows=124, skipfooter=21, usecols='P:Y', index_col= 0)
    df_19 = pd.read_excel(path_1, sheet_name= 'OVALAMA - GÜNDÜZ (2)', skiprows=134, skipfooter=11, usecols='B:K', index_col= 0)
    df_20 = pd.read_excel(path_1, sheet_name= 'OVALAMA - GÜNDÜZ (2)', skiprows=134, skipfooter=11, usecols='P:Y', index_col= 0)
    df_21 = pd.read_excel(path_1, sheet_name= 'OVALAMA - GECE (3)', skiprows=98, skipfooter=41, usecols='B:I', index_col= 0)
    df_22 = pd.read_excel(path_1, sheet_name= 'OVALAMA - GECE (3)', skiprows=108, skipfooter=31, usecols='B:I', index_col= 0)
    df_23 = pd.read_excel(path_1, sheet_name= 'OVALAMA - GECE (3)', skiprows=118, skipfooter=21, usecols='B:I', index_col= 0)
    df_24 = pd.read_excel(path_1, sheet_name= 'OVALAMA - GECE (3)', skiprows=98, skipfooter=41, usecols='O:X', index_col= 0)
    df_25 = pd.read_excel(path_1, sheet_name= 'OVALAMA - GECE (3)', skiprows=108, skipfooter=31, usecols='O:X', index_col= 0)
    df_26 = pd.read_excel(path_1, sheet_name= 'OVALAMA - GECE (3)', skiprows=118, skipfooter=21, usecols='O:X', index_col= 0)
    df_27 = pd.read_excel(path_1, sheet_name= 'OVALAMA - GECE (3)', skiprows=128, skipfooter=11, usecols='B:I', index_col= 0)
    df_28 = pd.read_excel(path_1, sheet_name= 'OVALAMA - GECE (3)', skiprows=128, skipfooter=11, usecols='O:U', index_col= 0)
    df_29 = pd.read_excel(path_1, sheet_name= 'MUV GÜNDÜZ', skiprows=50, skipfooter=0, usecols='E:L', index_col= 0)
    df_30 = pd.read_excel(path_1, sheet_name= 'MUV GECE', skiprows=50, skipfooter=0, usecols='E:L', index_col= 0)
    
    for table in range(1,31):
        for j in range(0,len(locals()[f'df_{table}'].loc['OPERATÖR'])):
            if pd.isna(locals()[f'df_{table}'].loc['OPERATÖR'][j]) == 0:
                if 'Unnamed:' not in locals()[f'df_{table}'].loc['OPERATÖR'][j] and 'TOPLAM' not in locals()[f'df_{table}'].loc['OPERATÖR'][j]:
                    dfmain.loc[dfmain.PERSONEL == locals()[f'df_{table}'].loc['OPERATÖR'][j].strip(), day] = locals()[f'df_{table}'].loc['ÜRETİM YÜZDESİ'][j]
        if pd.isnull(locals()[f'df_{table}'].index.name) == 0:
            if '&' in locals()[f'df_{table}'].index.name:
                names = locals()[f'df_{table}'].index.name.split('&')
                for name in names:
                    dfmain.loc[dfmain.PERSONEL == name.strip(), day] = locals()[f'df_{table}'].loc['ÜRETİM YÜZDESİ'][-1]
            else:
                dfmain.loc[dfmain.PERSONEL == locals()[f'df_{table}'].index.name.strip(), day] = locals()[f'df_{table}'].loc['ÜRETİM YÜZDESİ'][-1]
    return dfmain
    
def montaj(day,month,year,path_1,dfmain):
    df_1 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=3, skipfooter=299, usecols='N:O')
    df_2 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=25, skipfooter=274, usecols='N:O')
    df_3 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=50, skipfooter=255, usecols='N:O')
    df_4 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=69, skipfooter=236, usecols='N:O')
    df_5 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=88, skipfooter=227, usecols='N:O')
    df_6 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=97, skipfooter=218, usecols='N:O')
    df_7 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=106, skipfooter=206, usecols='N:O')
    df_8 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=121, skipfooter=186, usecols='N:O')
    df_9 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=141, skipfooter=182, usecols='N:O')
    df_10 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=147, skipfooter=176, usecols='N:O')
    df_11 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=155, skipfooter=165, usecols='N:O')
    df_12 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=159, skipfooter=161, usecols='N:O')
    df_13 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=163, skipfooter=157, usecols='N:O')
    df_14 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=167, skipfooter=153, usecols='N:O')
    df_15 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=171, skipfooter=149, usecols='N:O')
    df_16 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=175, skipfooter=145, usecols='N:O')
    df_17 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=179, skipfooter=141, usecols='N:O')
    df_18 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=183, skipfooter=137, usecols='N:O')
    df_19 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=187, skipfooter=133, usecols='N:O')
    df_20 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=191, skipfooter=129, usecols='N:O')
    df_21 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=195, skipfooter=118, usecols='N:O')
    df_22 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=210, skipfooter=113, usecols='N:O')
    df_22 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=213, skipfooter=110, usecols='N:O')
    df_23 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=216, skipfooter=107, usecols='N:O')
    df_24 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=219, skipfooter=104, usecols='N:O')
    df_25 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=222, skipfooter=94, usecols='N:O')
    df_26 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=231, skipfooter=92, usecols='N:O')
    df_27 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=234, skipfooter=89, usecols='N:O')
    df_28 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=237, skipfooter=86, usecols='N:O')
    df_29 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=240, skipfooter=83, usecols='N:O')
    df_30 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=244, skipfooter=79, usecols='N:O')
    df_31 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=247, skipfooter=76, usecols='N:O')
    df_32 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=250, skipfooter=73, usecols='N:O')
    df_33 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=254, skipfooter=69, usecols='N:O')
    df_34 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=257, skipfooter=66, usecols='N:O')
    df_35 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=260, skipfooter=59, usecols='N:O')
    df_36 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=267, skipfooter=56, usecols='N:O')
    df_37 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=270, skipfooter=53, usecols='N:O')
    df_38 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=273, skipfooter=50, usecols='N:O')
    df_39 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=276, skipfooter=47, usecols='N:O')
    df_40 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=279, skipfooter=44, usecols='N:O')
    df_41 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=282, skipfooter=41, usecols='N:O')
    df_42 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=288, skipfooter=35, usecols='N:O')
    df_43 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=291, skipfooter=32, usecols='N:O')
    df_44 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=294, skipfooter=29, usecols='N:O')
    df_45 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=300, skipfooter=23, usecols='N:O')
    df_46 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GÜNDÜZ', skiprows=303, skipfooter=20, usecols='N:O')
    df_47 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=3, skipfooter=265, usecols='N:O')
    df_48 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=21, skipfooter=247, usecols='N:O')
    df_49 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=39, skipfooter=228, usecols='N:O')
    df_50 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=58, skipfooter=202, usecols='N:O')
    df_51 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=84, skipfooter=192, usecols='N:O')
    df_52 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=94, skipfooter=182, usecols='N:O')
    df_53 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=104, skipfooter=171, usecols='N:O')
    df_54 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=118, skipfooter=155, usecols='N:O')
    df_55 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=135, skipfooter=150, usecols='N:O')
    df_56 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=141, skipfooter=144, usecols='N:O')
    df_57 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=146, skipfooter=136, usecols='N:O')
    df_58 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=150, skipfooter=132, usecols='N:O')
    df_59 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=154, skipfooter=128, usecols='N:O')
    df_60 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=158, skipfooter=124, usecols='N:O')
    df_61 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=162, skipfooter=121, usecols='N:O')
    df_62 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=165, skipfooter=118, usecols='N:O')
    df_63 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=168, skipfooter=107, usecols='N:O')
    df_64 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=182, skipfooter=103, usecols='N:O')
    df_65 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=186, skipfooter=99, usecols='N:O')
    df_66 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=189, skipfooter=96, usecols='N:O')
    df_67 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=193, skipfooter=92, usecols='N:O')
    df_68 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=196, skipfooter=89, usecols='N:O')
    df_69 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=203, skipfooter=82, usecols='N:O')
    df_70 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=206, skipfooter=79, usecols='N:O')
    df_71 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=209, skipfooter=76, usecols='N:O')
    df_72 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=212, skipfooter=73, usecols='N:O')
    df_73 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=216, skipfooter=69, usecols='N:O')
    df_74 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=219, skipfooter=66, usecols='N:O')
    df_75 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=222, skipfooter=63, usecols='N:O')
    df_76 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=225, skipfooter=60, usecols='N:O')
    df_77 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=232, skipfooter=53, usecols='N:O')
    df_78 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=235, skipfooter=50, usecols='N:O')
    df_79 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=238, skipfooter=47, usecols='N:O')
    df_80 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=241, skipfooter=44, usecols='N:O')
    df_81 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=244, skipfooter=41, usecols='N:O')
    df_82 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=250, skipfooter=35, usecols='N:O')
    df_83 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=253, skipfooter=32, usecols='N:O')
    df_84 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=256, skipfooter=29, usecols='N:O')
    df_85 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=262, skipfooter=23, usecols='N:O')
    df_86 = pd.read_excel(path_1, sheet_name= f'{day}.{month}.{year} GECE', skiprows=265, skipfooter=20, usecols='N:O')


    gday = f'G{day}'
    for table in range(1,87):
        for j in range(0,len(locals()[f'df_{table}'].iloc[:,1])):
            if pd.isna(locals()[f'df_{table}'].iloc[j,1]) == 0:
                dfmain.loc[dfmain.PERSONEL == locals()[f'df_{table}'].iloc[j,1].strip(), gday] = locals()[f'df_{table}'].iloc[0,0]
    return dfmain