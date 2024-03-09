# app.py - главный файл приложения

import streamlit as st

st.markdown("<h1 style='text-align: center;color:orange;'>Вас приветствует</h1>", unsafe_allow_html=True)

st.markdown("<h2 style='text-align: center;color:orange;'>мастер отчетов</h2>", unsafe_allow_html=True)
st.markdown("<h2 style='text-align: center;color:orange;'>независимой оценки качества услуг</h2>", unsafe_allow_html=True)

st.subheader('перейдите на страницу с требуемым расчетом', divider='red')

st.subheader(':red[перед загрузкой файлов убедитесь что они соответствуют образцу]:', divider='red')
st.subheader('образцы можно посмотреть в меню слева - образцы для загрузочных файлов')

if st.button("**Расчет для организаций культуры**"):
    st.switch_page("pages/Для_организаций_культуры.py")
if st.button("**Расчет для социальных организаций**"):
    st.switch_page("pages/Для_социальных_организаций.py")
if st.button("**Расчет для медицинских организаций**"):
    st.switch_page("pages/Для_медицинских_организаций.py")
       
