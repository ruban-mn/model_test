#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# app.py - главный файл приложения

import streamlit as st

def main():
    st.title("Independent Quality of Service Assessment Report Master")
    
    session_state = st.session_state
    if 'button_pressed' not in session_state:
        session_state.button_pressed = False

    menu = ["Home", "Contact"]
    choice = st.sidebar.selectbox("Menu", menu)

    if choice == "Home":
        st.write("Welcome to the Independent Quality of Service Assessment Report Master")
        st.subheader('Go to the page with the required calculation', divider='rainbow')
       
        if st.button('Calculations for Cultural Organizations'):
            session_state.button_pressed = True
            session_state.selected_option = 'organization_culture'
        
        if st.button('Calculations for Social Organizations'):
            session_state.button_pressed = True
            session_state.selected_option = 'social_organization'

        if st.button('Calculations for Medical Organizations'):
            session_state.button_pressed = True
            session_state.selected_option = 'medical_organization'
    
    elif choice == "Contact":
        st.write("Page with contact information.")
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

