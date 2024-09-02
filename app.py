import streamlit as st

st.title('나의 첫 Streamlit 앱')
st.header('이것은 헤더입니다')
st.write('Streamlit을 사용하면 이렇게 간단하게 웹 앱을 만들 수 있습니다!')

# 사용자 입력 받기
user_input = st.text_input('이름을 입력해 주세요:', '')

if user_input:
    st.write('안녕하세요, ', user_input, '님!')
