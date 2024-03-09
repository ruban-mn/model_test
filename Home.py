#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# app.py - главный файл приложения

import streamlit as st

st.header("Вас приветствует мастер отчетов независимой оценки качества услуг")
st.subheader('перейдите на страницу с требуемым расчетом', divider='rainbow')
       
if st.button('Расчеты для организаций культуры'):
    with open('pages/myself_an_analitik_8_2.py', 'r') as file:
        code = file.read()
    exec(code)
