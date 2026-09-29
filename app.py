import streamlit as st
import streamlit.components.v1 as components
import os

# Cấu hình trang Streamlit full chiều rộng
st.set_page_config(
    page_title="Dashboard INF TQG - Executive Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Thêm CSS ẩn Header/Footer mặc định của Streamlit để giao diện đẹp tuyệt đối
st.markdown("""
    <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        .block-container {
            padding-top: 0rem !important;
            padding-bottom: 0rem !important;
            padding-left: 0rem !important;
            padding-right: 0rem !important;
            max-width: 100% !important;
        }
        iframe {
            border: none;
        }
    </style>
""", unsafe_allow_html=True)

# Đọc file HTML giao diện
def load_html():
    html_path = os.path.join(os.path.dirname(__file__), "dashboard.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return f.read()

html_content = load_html()

# Hiển thị giao diện HTML/JS nguyên bản
components.html(html_content, height=1000, scrolling=True)
