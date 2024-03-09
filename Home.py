# app.py - главный файл приложения

import streamlit as st

st.markdown("<h1 style='text-align: center;'>Текст выровнен по центру</h1>", unsafe_allow_html=True)

st.title("Вас приветствует")
st.header(":orange[мастер отчетов независимой оценки качества услуг]")
st.subheader(':blue[перейдите на страницу с требуемым расчетом]:', divider='rainbow')

st.markdown('<button style="font-weight: bold;">Расчет для организаций культуры</button>', unsafe_allow_html=True)
if st.button("Расчет для организаций культуры"):
    st.switch_page("pages/Для_организаций_культуры.py")
if st.button("Расчет для социальных организаций"):
    st.switch_page("pages/Для_социальных_организаций.py")
if st.button("Расчет для медицинских организаций"):
    st.switch_page("pages/Для_медицинских_организаций.py")
       
