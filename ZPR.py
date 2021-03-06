from ast import main
import pandas as pd
import numpy as np
import math

def anal_sklad(realiz_avto):
    #выводим все авто проданные за 2 месяца
    kol_prod_avto=[]  
    for i in range(realiz_avto.shape[0]):
        if '.06' in realiz_avto.iloc[i]['Дата'] or '.05' in realiz_avto.iloc[i]['Дата'] :
            kol_prod_avto.append(realiz_avto.iloc[i]['Дата'])

    return(len(kol_prod_avto)/2)  

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

def skok_avto_VR(sklad_avto): 
    schetchik = 0
    #ищем колличество новых авто в файле
    for i in range(sklad_avto.shape[0]):
        if sklad_avto.iloc[i]["№ п/п"] == "№ п/п":
            schetchik+=1
            if schetchik == 1:
                return(sklad_avto.iloc[i-5]["№ п/п"])

def skok_avto_FD(sklad_avto): 
    #ищем колличество новых авто в файле
    schetchik = 0
    for i in range(sklad_avto.shape[0]):
        if sklad_avto.iloc[i]["№ п/п"] == "№ п/п":
            schetchik+=1
            if schetchik == 1:
                return(sklad_avto.iloc[i-4]["№ п/п"])   

def skok_avto_o(sklad_avto):
    schetchik = 0
    #ищем колличество новых авто в файле
    for i in range(sklad_avto.shape[0]):
        if sklad_avto.iloc[i]["№ п/п"] == "№ п/п":
            schetchik+=1
            if schetchik == 1:
                return(sklad_avto.iloc[i-5]["№ п/п"]) 

def skok_avto_G(sklad_avto):
    schetchik = 0
    #ищем колличество новых авто в файле
    for i in range(sklad_avto.shape[0]):
        if sklad_avto.iloc[i]["№ п/п"] == "№ п/п":
            schetchik+=1
            if schetchik == 2:
                return(sklad_avto.iloc[i-2]["№ п/п"]) 

def skok_avto_S(sklad_avto):
    schetchik = 0
    #ищем колличество новых авто в файле
    for i in range(sklad_avto.shape[0]):
        if sklad_avto.iloc[i]["№ п/п"] == "№ п/п":
            schetchik+=1
            if schetchik == 1:
                return(sklad_avto.iloc[i-4]["№ п/п"]) 
def vr_motors_date():

    #подгрузка данных 
    other_metrick = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(VR-Motors).xlsx',
                                    sheet_name = 'Прочие метрики', usecols='A:N')
    sklad_avto = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(VR-Motors).xlsx',
                                    sheet_name = 'Авто на складах', skiprows=4, usecols='A:S')
    realiz_avto = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(VR-Motors).xlsx',
                                    sheet_name = 'Реализация авто за полгода')
    money_on_schet = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(VR-Motors).xlsx',
                                    sheet_name = 'Денег на счетах', skiprows=1, usecols='A:S')
    plan_spend = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(VR-Motors).xlsx',
                                    sheet_name = 'Планируемые затраты')


    if other_metrick.iloc[0]['значения'] == 0:
        new_avto_kol=skok_avto_VR(sklad_avto)
    else:
        print("Error")
    vazhnost_pokupki=anal_vazh_sklad(realiz_avto,new_avto_kol,other_metrick)
    print("Заказан автовоз? Введите Да или Нет")
    otvet_avtovoz = input()
    if otvet_avtovoz.upper() == "ДА":
        print("Сколько авто купили")
        new_avto_kol=int(new_avto_kol)+int(input())
        vazhnost_pokupki=anal_vazh_sklad(realiz_avto,new_avto_kol,other_metrick)
        print(vazhnost_pokupki)
    else:
        print(vazhnost_pokupki)

def fd_motors_date():
    #подгрузка данных
    other_metrick = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(FD-Motors).xlsx',
                                    sheet_name = 'Прочие метрики', usecols='A:N')
    sklad_avto = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(FD-Motors).xlsx',
                                    sheet_name = 'Авто на складах', skiprows=4, usecols='A:S')
    realiz_avto = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(FD-Motors).xlsx',
                                    sheet_name = 'Реализация авто за полгода')
    money_on_schet = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(FD-Motors).xlsx',
                                    sheet_name = 'Денег на счетах', skiprows=1, usecols='A:S')
    plan_spend = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(FD-Motors).xlsx',
                                    sheet_name = 'Планируемые затраты', usecols='A:I')

    if other_metrick.iloc[0]['значения'] == 0:
        new_avto_kol=skok_avto_FD(sklad_avto)
    else:
        print("Error")
    vazhnost_pokupki=anal_vazh_sklad(realiz_avto,new_avto_kol,other_metrick)
    print("Заказан автовоз? Введите Да или Нет")
    otvet_avtovoz = input()
    if otvet_avtovoz.upper() == "ДА":
        print("Сколько авто купили")
        new_avto_kol=int(new_avto_kol)+int(input())
        vazhnost_pokupki=anal_vazh_sklad(realiz_avto,new_avto_kol,other_metrick)
        print(vazhnost_pokupki)
    else:
        print(vazhnost_pokupki)  

def o_motors_date():
    
    #подгрузка данных
    other_metrick = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(O-Motors).xlsx',
                                    sheet_name = 'Прочие метрики', usecols='A:N')
    sklad_avto = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(O-Motors).xlsx',
                                    sheet_name = 'Авто на складах', skiprows=4, usecols='A:S')
    realiz_avto = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(O-Motors).xlsx',
                                    sheet_name = 'Реализация авто за полгода')
    money_on_schet = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(O-Motors).xlsx',
                                    sheet_name = 'Денег на счетах', skiprows=1, usecols='A:S')
    plan_spend = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(O-Motors).xlsx',
                                    sheet_name = 'Планируемые затраты', usecols='A:I')
    

    #Основное тело алгоритма проверки на значимость
    if other_metrick.iloc[0]['значения'] == 0:
        new_avto_kol=skok_avto_o(sklad_avto)
    else:
        print("Error")
    # оценка относительно склада
    vazhnost_pokupki=anal_vazh_sklad(realiz_avto, new_avto_kol, other_metrick)
    print("Заказан автовоз? Введите Да или Нет")
    otvet_avtovoz = input()
    if otvet_avtovoz.upper() == "ДА":
        print("Сколько авто купили")
        new_avto_kol=int(new_avto_kol)+int(input())
        vazhnost_pokupki=anal_vazh_sklad(realiz_avto, new_avto_kol, other_metrick)
        print(vazhnost_pokupki)
    else:
        print(vazhnost_pokupki)
 
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
    vazhnost_pokupki=anal_vazh_sklad(realiz_avto, new_avto_kol, other_metrick)
    print("Заказан автовоз? Введите Да или Нет")
    otvet_avtovoz = input()
    if otvet_avtovoz.upper() == "ДА":
        print("Сколько авто купили")
        new_avto_kol=int(new_avto_kol)+int(input())
        vazhnost_pokupki=anal_vazh_sklad(realiz_avto, new_avto_kol, other_metrick)
        print(vazhnost_pokupki)
    else:
        print(vazhnost_pokupki)

def s_motors_date():
    
    #подгрузка данных
    other_metrick = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(S-Motors).xlsx',
                                    sheet_name = 'Прочие метрики', usecols='A:N')
    sklad_avto = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(S-Motors).xlsx',
                                    sheet_name = 'Авто на складах', skiprows=4, usecols='A:S')
    realiz_avto = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(S-Motors).xlsx',
                                    sheet_name = 'Реализация авто за полгода')
    money_on_schet = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(S-Motors).xlsx',
                                    sheet_name = 'Денег на счетах', skiprows=1, usecols='A:S')
    plan_spend = pd.read_excel('C:\\Users\\Artem\\Desktop\\Coding\\Оценочная модель\\bay_avto(S-Motors).xlsx',
                                    sheet_name = 'Планируемые затраты', usecols='A:I')
    

    #Основное тело алгоритма проверки на значимость
    if other_metrick.iloc[0]['значения'] == 0:
        new_avto_kol=skok_avto_S(sklad_avto)
    else:
        print("Error")
    # оценка относительно склада
    vazhnost_pokupki=anal_vazh_sklad(realiz_avto, new_avto_kol, other_metrick)
    print("Заказан автовоз? Введите Да или Нет")
    otvet_avtovoz = input()
    if otvet_avtovoz.upper() == "ДА":
        print("Сколько авто купили")
        new_avto_kol=int(new_avto_kol)+int(input())
        vazhnost_pokupki=anal_vazh_sklad(realiz_avto, new_avto_kol, other_metrick)
        print(vazhnost_pokupki)
    else:
        print(vazhnost_pokupki)

def main():
    print("Введите компанию: 1 - ВР-моторс, 2 - ВР-Авто, 3 - ВР-Октава, 4 - А-Моторс, 5 - ВР-Сакура")
    compania = input()
    if compania == "1":
        vr_motors_date()
    elif compania == "2":
        fd_motors_date()
    elif compania == "3":
        o_motors_date()
    elif compania == "4":
        a_motors_date()
    elif compania == "5":
        s_motors_date()

if __name__ == '__main__':
    main()
