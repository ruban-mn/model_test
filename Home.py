#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# app.py - главный файл приложения

import streamlit as st

def main():
    st.title("Мастер отчетов независимой оценки качества услуг")

    menu = ["Главная", "Контакты"]
    choice = st.sidebar.selectbox("Меню", menu)

    if choice == "Главная":
        st.subheader("Вас приветствует мастер отчетов независимой оценки качества услуг")
        st.subheader('перейдите на страницу с требуемым расчетом', divider='rainbow')
       
        if st.button('Расчеты для организаций культуры'):
            with open('https://github.com/ruban-mn/model1/blob/pages/myself_an_analitik_8_2.py') as file:
                code = file.read()
            exec(code)
        
        if st.button('Расчеты для социальных организаций'):
            with open('./pages/myself_an_analitik_8_3.py', 'r') as file:
                code = file.read()
            exec(code)

        if st.button('Расчеты для медицинских организаций'):
            with open('./pages/myself_an_analitik_8_4.py', 'r') as file:
                code = file.read()
            exec(code)

    elif choice == "Контакты":
        st.write("Страница с контактной информацией.")
        # Вставьте здесь содержимое для страницы 'Контакты'

if __name__ == '__main__':
    main()

