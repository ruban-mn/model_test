#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# app.py - главный файл приложения

import streamlit as st

def main():
    st.header("Мастер отчетов независимой оценки качества услуг")
    
    session_state = st.session_state
    if 'button_pressed' not in session_state:
        session_state.button_pressed = False

    menu = ["Главная", "Контакты"]
    choice = st.sidebar.selectbox("Menu", menu)

    if choice == "Главная":
        st.header("Вас приветствует мастер отчетов независимой оценки качества услуг")
        st.subheader('перейдите на страницу с требуемым расчетом', divider='rainbow')
       
        if st.button('Расчеты для организаций культуры'):
            session_state.button_pressed = True
            session_state.selected_option = 'organization_culture'
        
        if st.button('Расчеты для социальных организаций'):
            session_state.button_pressed = True
            session_state.selected_option = 'social_organization'

        if st.button('Расчеты для медицинских организаций'):
            session_state.button_pressed = True
            session_state.selected_option = 'medical_organization'
    
    elif choice == "Контакты":
        st.write("Разработчик Рубан Мария Николаевна, Ma77ruban@yandex.ru")
        # Add content for the 'Contact' page here

    if session_state.button_pressed:
        if session_state.selected_option == 'organization_culture':
            with open('myself_an_analitik_8_2.py', 'r') as file:
                code = file.read()
            exec(code)
        
        if session_state.selected_option == 'social_organization':
            with open('myself_an_analitik_8_3.py', 'r') as file:
                code = file.read()
            exec(code)

        if session_state.selected_option == 'medical_organization':
            with open('myself_an_analitik_8_4.py', 'r') as file:
                code = file.read()
            exec(code)

if __name__ == '__main__':
    main()

