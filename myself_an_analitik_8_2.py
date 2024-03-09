#!/usr/bin/env python
# coding: utf-8


import pandas as pd
import numpy as np
import openpyxl
import docx
import streamlit as st
import io
import time as tm


# In[7]:


st.subheader('Независимой оценка качества услуг организаций культуры')


# In[6]:


uploaded_file = st.file_uploader("Загрузите файл сводную по чек-листам", type=["xls", "xlsx"])

if uploaded_file is not None:
    # Чтение данных из загруженного файла Excel
    chek_list = pd.read_excel(uploaded_file)


# In[7]:


uploaded_file = st.file_uploader("Загрузите файл с массивом анкет", type=["xls", "xlsx"])

if uploaded_file is not None:
    # Чтение данных из загруженного файла Excel
    Answers_respond = pd.read_excel(uploaded_file)


# In[8]:


## подгружаем сводную по чек листам
#chek_list=pd.read_excel(r"C:\Users\user\чек лист гулькевичи.xlsx")
##подгружаем массив с ответами респондентов
#Answers_respond=pd.read_excel(r"C:\Users\user\Анкета Гулькевичи НОК культура (Ответы).xlsx")


# In[ ]:


tm.sleep(30)


# In[9]:


Answers_respond_list = Answers_respond.columns.tolist() ##извлекаем наименования столбцов в список


# In[ ]:


New_col = []  # Создаем пустой список
for i in range(18):  # Цикл от 0 до 18
    sim = i   # присваиваем номер
    New_col.append('v' + str(sim))  # добавляем новый номер вопрса в список


# In[ ]:


dictionary = dict(zip(Answers_respond_list, New_col)) # создаем  словарь для переименования стобцов


# In[ ]:


Answers_respond = Answers_respond.rename(columns=dictionary) # переименовываем столбцы в начальном датафрейме


# In[ ]:


# Создание нового DataFrame для хранения результатов подсчета, считам количество ответов да на вопросы анкеты
ans_res = pd.DataFrame({'v0': Answers_respond['v0'].unique()})

# Используем цикл для подсчета значений и создания новых столбцов
for col in New_col:
    value = 'Да'  # Значение, которое мы считаем
    count_col_name = f'_{col}_'
    counts = Answers_respond[Answers_respond[col] == value].groupby('v0').size().reset_index(name=count_col_name)
    ans_res = ans_res.merge(counts, on='v0', how='left')


# In[ ]:


ans_res = ans_res.dropna(axis=1) # Удаляем столбцы со значением NaN
ans_res['v0'] = ans_res['v0'].str.replace('.', '')# Удаляем точку из наименований организаций
ans_res = ans_res.sort_values(by='v0') # сортируем таблицу по возрастанию по столбцу наименования
ans_res = ans_res.reset_index(drop=True)


# In[ ]:


col_ob = Answers_respond.groupby('v0').size().reset_index(name='Ч_общ')
col_ob['v0'] = col_ob['v0'].str.replace('.', '')# Удаляем точку из наименований организаций
col_ob = col_ob.sort_values(by='v0') # сортируем таблицу по возрастанию по столбцу наименования
col_ob = col_ob.reset_index(drop=True)
all_ans = col_ob['Ч_общ']


# In[ ]:


name_org = chek_list.filter(like='Наименование организации').copy()


# In[10]:


Raschet_ballov = name_org
Raschet_ballov['Истенд'] = chek_list.filter(like='На СТЕНДЕ').sum(axis=1)
Raschet_ballov['Исайт'] = chek_list.filter(like='На САЙТЕ').sum(axis=1)
Raschet_ballov['Инорм-стенд'] = chek_list.filter(like='На СТЕНДЕ').count(axis=1)
Raschet_ballov['Инорм-сайт'] = chek_list.filter(like='На САЙТЕ').count(axis=1)
Raschet_ballov['Пинф'] = round(0.5*((Raschet_ballov['Истенд']/Raschet_ballov['Инорм-стенд'])+(Raschet_ballov['Исайт']/Raschet_ballov['Инорм-сайт']))*100, 2)
Raschet_ballov['Тдист'] = 30
Raschet_ballov['Сдист'] = chek_list.filter(like='Наличие и функционирование на официальном сайте').sum(axis=1)
Raschet_ballov['Пдист'] = Raschet_ballov['Тдист']*Raschet_ballov['Сдист']
Raschet_ballov['Пдист'].where(Raschet_ballov['Пдист'] <= 100, 100, inplace=True)
Raschet_ballov['Устенд'] = ans_res['_v2_']
Raschet_ballov['Усайт'] = ans_res['_v4_']
Raschet_ballov['Уобщ-стенд'] = ans_res['_v1_']
Raschet_ballov['Уобщ-сайт'] = ans_res['_v3_']
Raschet_ballov['Поткруд'] = round(0.5*((Raschet_ballov['Устенд']/Raschet_ballov['Уобщ-стенд'])+(Raschet_ballov['Усайт']/Raschet_ballov['Уобщ-сайт']))*100, 2)
Raschet_ballov['К1'] = round(0.3*Raschet_ballov['Пинф'] + 0.3*Raschet_ballov['Пдист'] + 0.4*Raschet_ballov['Поткруд'], 2)
Raschet_ballov['Ткомф'] = chek_list.filter(like='Обеспечение в организации комфортных условий').sum(axis=1)
Raschet_ballov['Скомф'] = 20
Raschet_ballov['Пкомф.усл'] = Raschet_ballov['Ткомф']*Raschet_ballov['Скомф']
Raschet_ballov['Пкомф.усл'].where(Raschet_ballov['Пкомф.усл'] <= 100, 100, inplace=True)
Raschet_ballov['Укомф'] = ans_res['_v5_']
Raschet_ballov['Чобщ'] = col_ob['Ч_общ']
Raschet_ballov['Пкомфуд'] = round(Raschet_ballov['Укомф']/Raschet_ballov['Чобщ']*100, 2)
Raschet_ballov['К2'] = round(0.5*Raschet_ballov['Пкомф.усл'] + 0.5*Raschet_ballov['Пкомфуд'], 2)
Raschet_ballov['Торгдост'] = chek_list.filter(like='Оборудование территории').sum(axis=1)
Raschet_ballov['Соргдост'] = 20
Raschet_ballov['Поргдост'] = Raschet_ballov['Торгдост']*Raschet_ballov['Соргдост']
Raschet_ballov['Туслугдост'] = chek_list.filter(like='Обеспечение в организации условий доступности').sum(axis=1)
Raschet_ballov['Суслугдост'] = 20
Raschet_ballov['Пуслугдост'] = Raschet_ballov['Туслугдост']*Raschet_ballov['Суслугдост']
Raschet_ballov['Пуслугдост'].where(Raschet_ballov['Пуслугдост'] <= 100, 100, inplace=True)
Raschet_ballov['Удост'] = ans_res['_v7_']
Raschet_ballov['Чинв'] = ans_res['_v6_']
Raschet_ballov['Пдостуд'] = round(Raschet_ballov['Удост']/Raschet_ballov['Чинв']*100, 2)
Raschet_ballov['К3'] = round(0.3*Raschet_ballov['Поргдост'] + 0.4*Raschet_ballov['Пуслугдост'] + 0.3*Raschet_ballov['Пдостуд'], 2)
Raschet_ballov['Уперв.конт'] = ans_res['_v8_']
Raschet_ballov['Чобщ1'] = all_ans
Raschet_ballov['Пперв.контуд'] = round(Raschet_ballov['Уперв.конт']/Raschet_ballov['Чобщ']*100, 2)
Raschet_ballov['Уоказ.услуг'] = ans_res['_v9_']
Raschet_ballov['Чобщ2'] = all_ans
Raschet_ballov['Показ.услугуд'] = round(Raschet_ballov['Уоказ.услуг']/Raschet_ballov['Чобщ']*100, 2)
Raschet_ballov['Увежл.дист'] = ans_res['_v11_']
Raschet_ballov['Чобщ_ус'] = ans_res['_v10_']
Raschet_ballov['Пвежл.дистуд'] = round(Raschet_ballov['Увежл.дист']/Raschet_ballov['Чобщ_ус']*100, 2)
Raschet_ballov['К4'] = round(0.4*Raschet_ballov['Пперв.контуд'] + 0.4*Raschet_ballov['Показ.услугуд'] + 0.2*Raschet_ballov['Пвежл.дистуд'], 2)
Raschet_ballov['Уреком'] = ans_res['_v12_']
Raschet_ballov['Чобщ3'] = all_ans
Raschet_ballov['Преком'] = round(Raschet_ballov['Уреком']/Raschet_ballov['Чобщ']*100, 2)
Raschet_ballov['Уорг.усл'] = ans_res['_v13_']
Raschet_ballov['Чобщ4'] = all_ans
Raschet_ballov['Порг.услуд'] = round(Raschet_ballov['Уорг.усл']/Raschet_ballov['Чобщ']*100, 2)
Raschet_ballov['Ууд'] = ans_res['_v14_']
Raschet_ballov['Чобщ5'] = all_ans
Raschet_ballov['Пуд'] = round(Raschet_ballov['Ууд']/Raschet_ballov['Чобщ']*100, 2)
Raschet_ballov['К5'] = round(0.3*Raschet_ballov['Преком'] + 0.2*Raschet_ballov['Порг.услуд'] + 0.5*Raschet_ballov['Пуд'], 2)
Raschet_ballov['Общий балл'] = round((Raschet_ballov['К1']+Raschet_ballov['К2']+Raschet_ballov['К3']+Raschet_ballov['К4']+Raschet_ballov['К5'])/5, 2)


# In[ ]:


button = st.button("получить готовый файл расчет баллов")

if button:
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    Raschet_ballov.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.close()
    output.seek(0)
    st.download_button(
        label="Загрузить",
        data=output,
        file_name='результаты.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


# In[ ]:




