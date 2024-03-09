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
        st.write("Вас приветствует мастер отчетов независимой оценки качества услуг")
        st.subheader('перейдите на страницу с требуемым расчетом', divider='rainbow')

        st.page_link('https://github.com/ruban-mn/model1/blob/pages/myself_an_analitik_8_2.py', label='Расчеты для организаций культуры')
        st.page_link('https://github.com/ruban-mn/model1/blob/pages/myself_an_analitik_8_3.py', label='Расчеты для социальных организаций')
        st.page_link('https://github.com/ruban-mn/model1/blob/pages/myself_an_analitik_8_4.py', label='Расчеты для медицинских организаций')

    elif choice == "Контакты":
        st.write("Страница с контактной информацией.")
        # Вставьте здесь содержимое для страницы 'Контакты'

if __name__ == '__main__':
    main()

