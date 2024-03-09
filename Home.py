# app.py - главный файл приложения

import streamlit as st

st.title("Вас приветствует")
st.header(":orange[мастер отчетов независимой оценки качества услуг]")
st.subheader(':blue[перейдите на страницу с требуемым расчетом]:', divider='rainbow')

if st.button("Расчет для организаций культуры"):
    st.switch_page("pages/Для_организаций_культуры.py")
if st.button("Расчет для социальных организаций"):
    st.switch_page("pages/Для_социальных_организаций.py")
if st.button("Расчет для медицинских организаций"):
    st.switch_page("pages/Для_медицинских_организаций.py")
       
