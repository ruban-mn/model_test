#!/usr/bin/env python
# coding: utf-8

# app.py - главный файл приложения

import streamlit as st

st.header("Вас приветствует мастер отчетов независимой оценки качества услуг")
st.subheader('перейдите на страницу с требуемым расчетом', divider='rainbow')
       
st.page_link("Home.py", label="Главная", icon="🏠")
st.page_link("pages/myself_an_analitik_8_2.py", label="Расчет для организаций культуры", icon="1️⃣")
st.page_link("pages/myself_an_analitik_8_3.py", label="Расчет для социальных организаций", icon="2️⃣")
st.page_link("pages/myself_an_analitik_8_4.py", label="Расчет для социальных организаций", icon="2️⃣")
