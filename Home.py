# app.py - главный файл приложения

import streamlit as st

st.markdown("<h1 style='text-align: center;'>Вас приветствует</h1>", unsafe_allow_html=True)

st.markdown("<h2 style='text-align: center;'>:orange[мастер отчетов]</h2>", unsafe_allow_html=True)

st.header(":orange[мастер отчетов]")
st.header(":orange[независимой оценки качества услуг]")
st.subheader(':blue[перейдите на страницу с требуемым расчетом]:', divider='rainbow')

st.subheader(':red[перед загрузкой файлов удетитесь что они соответствуют образцу]:', divider='rainbow')
st.subheader('образцы можно посмотреть в меню слева - образцы для загрузочных файлов', divider='rainbow')

if st.button("Bold[Расчет для организаций культуры]"):
    st.switch_page("pages/Для_организаций_культуры.py")
if st.button("Расчет для социальных организаций"):
    st.switch_page("pages/Для_социальных_организаций.py")
if st.button("Расчет для медицинских организаций"):
    st.switch_page("pages/Для_медицинских_организаций.py")
       
