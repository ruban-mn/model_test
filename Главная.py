# app.py - главный файл приложения

import streamlit as st

st.header("Вас приветствует мастер отчетов независимой оценки качества услуг")
st.subheader(':blue[перейдите на страницу с требуемым расчетом]:', divider='rainbow')

if st.button("Расчет для организаций культуры"):
    st.switch_page("pages/Для организаций культуры.py")
if st.button("Расчет для социальных организаций"):
    st.switch_page("pages/Для_социальных_организаций.py")
if st.button("Расчет для медицинских организаций"):
    st.switch_page("pages/myself_an_analitik_8_4.py")
       
