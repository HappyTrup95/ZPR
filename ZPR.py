from ast import main
from cmath import nan
import pandas as pd
import numpy as np
import math

def anal_vazh_sklad(realiz_avto,new_avto_kol,other_metrick):
    # оценка относительно склада
    prodazh_avto=math.ceil(anal_sklad(realiz_avto))
    real_1_avto_t = 30/(prodazh_avto)
    ostalos_na_zakupku = (new_avto_kol-1)*real_1_avto_t
    avto_v_puty=other_metrick.iloc[1][1]
    vazh_zakupki= int(ostalos_na_zakupku)/int(avto_v_puty)
    if vazh_zakupki >= 1:
        vazh_zakupki = 1/vazh_zakupki
    else:
        vazh_zakupki=1/vazh_zakupki
    return(vazh_zakupki)  


def anal_sklad(realiz_avto):
    kol_prod_avto=[]
    for i in realiz_avto.items():
        if '.06' in realiz_avto.iloc[i]['Дата'] or '.05' in realiz_avto.iloc[i]['Дата'] :
            kol_prod_avto.append(realiz_avto.iloc[i]['Дата'])

    return(len(kol_prod_avto)/2)
   
   

def skok_avto_G(sklad_avto):
    schetchik = 0
    #ищем колличество новых авто в файле
    for i in range(100):
        if sklad_avto.iloc[i]["№ п/п"] == "№ п/п":
            schetchik+=1
            if schetchik == 2:
                return(sklad_avto.iloc[i-2]["№ п/п"])  


def a_motors_date():
    
    #подгрузка данных
    other_metrick = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(A-Motors).xlsx',
                                    sheet_name = 'Прочие метрики', usecols='A:N')
    sklad_avto = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(A-Motors).xlsx',
                                    sheet_name = 'Авто на складах', skiprows=4, usecols='A:S')
    realiz_avto = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(A-Motors).xlsx',
                                    sheet_name = 'Реализация авто за полгода')
    money_on_schet = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(A-Motors).xlsx',
                                    sheet_name = 'Денег на счетах', skiprows=1, usecols='A:S')
    plan_spend = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(A-Motors).xlsx',
                                    sheet_name = 'Планируемые затраты', usecols='A:I')
    

    #Основное тело алгоритма проверки на значимость
    if other_metrick.iloc[0]['значения'] == 0:
        new_avto_kol=skok_avto_G(sklad_avto)
    else:
        print("Error")
    # оценка относительно склада
    vazhnost_pokupki=anal_vazh_sklad(realiz_avto,new_avto_kol,other_metrick)
    print(vazhnost_pokupki)



def main():
    print("Введите компанию 1 - ВР-моторс, 2 - ВР-Авто, 3 - ВР-Октава, 4 - А-Моторс, 5 - ВР-Сакура")
    compania = input()

    if compania == "4":
        a_motors_date()

if __name__ == '__main__':
    main()