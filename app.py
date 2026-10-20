import io
import os
import time
import pandas as pd
import requests
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# 1. CẤU HÌNH TRANG & GIAO DIỆN CHUẨN CỔNG THÔNG TIN TRƯỜNG THCS CHÍ CÔNG
# ==============================================================================
st.set_page_config(
    page_title="Hệ Thống Thi Đua Lớp 8B3 - THCS Chí Công",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Thay đường dẫn link web app /exec từ Google Apps Script của bạn vào đây nếu có
API_URL = "https://script.google.com/macros/s/AKfycbwoy6GT8XEkwwrLTSXrIk3tMj6lzATeiwrw02cWYz7yEXDhzEtDOoyL8VssgXuJE08q/exec"
DATA_FILE = "du_lieu_thi_dua.csv"

# ==============================================================================
# 2. BỘ CSS GIAO DIỆN HIỆN ĐẠI (SIDEBAR THẺ BO TRÒN & BANNER TRƯỜNG HỌC)
# ==============================================================================
st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Segoe+UI:wght@400;600;700&display=swap');
        html, body, [class*="css"] {
            font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
        }
        .main {
            background-color: #f4f6f9;
        }

        /* Nền và khung Sidebar */
        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #f8fafc 0%, #edf2f7 100%) !important;
            border-right: 1px solid #e2e8f0;
            box-shadow: 3px 0 12px rgba(0,0,0,0.03);
        }
        
        /* Card nhận diện thương hiệu trường trong Sidebar */
        .sidebar-brand-card {
            background: #ffffff;
            border-radius: 12px;
            padding: 14px 16px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 2px 6px rgba(11, 79, 138, 0.08);
            margin-bottom: 18px;
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .sidebar-brand-icon {
            width: 44px;
            height: 44px;
            background: #0b4f8a;
            color: #ffffff;
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.4rem;
        }
        .sidebar-brand-title {
            font-size: 1.05rem;
            font-weight: 800;
            color: #0b4f8a;
            line-height: 1.2;
        }
        .sidebar-brand-sub {
            font-size: 0.78rem;
            color: #64748b;
            font-weight: 500;
        }

        /* Menu điều hướng dạng thẻ bo tròn (ẩn radio tròn trơn) */
        div[data-testid="stRadio"] > div {
            gap: 8px !important;
        }
        div[data-testid="stRadio"] label {
            background: #ffffff !important;
            border: 1px solid #e2e8f0 !important;
            border-radius: 10px !important;
            padding: 10px 14px !important;
            margin-bottom: 2px !important;
            cursor: pointer !important;
            transition: all 0.2s ease !important;
            box-shadow: 0 1px 3px rgba(0,0,0,0.02) !important;
            display: flex !important;
            align-items: center !important;
        }
        div[data-testid="stRadio"] label > div:first-child {
            display: none !important;
        }
        div[data-testid="stRadio"] label p {
            font-size: 0.92rem !important;
            font-weight: 600 !important;
            color: #334155 !important;
            margin: 0 !important;
        }
        div[data-testid="stRadio"] label:hover {
            background-color: #f0f7ff !important;
            border-color: #93c5fd !important;
            transform: translateX(4px);
        }
        div[data-testid="stRadio"] label[data-checked="true"],
        div[data-testid="stRadio"] label:has(input:checked) {
            background: #0b4f8a !important;
            border-color: #0b4f8a !important;
            box-shadow: 0 3px 8px rgba(11, 79, 138, 0.25) !important;
        }
        div[data-testid="stRadio"] label[data-checked="true"] p,
        div[data-testid="stRadio"] label:has(input:checked) p {
            color: #ffffff !important;
            font-weight: 700 !important;
        }

        .sidebar-section-label {
            font-size: 0.75rem;
            text-transform: uppercase;
            letter-spacing: 0.6px;
            font-weight: 700;
            color: #64748b;
            margin: 14px 0 6px 4px;
            display: flex;
            align-items: center;
            gap: 6px;
        }

        /* Banner trường học ở đầu nội dung */
        .school-top-bar {
            background: linear-gradient(90deg, #0b4f8a 0%, #1565c0 100%);
            color: #ffffff;
            padding: 12px 20px;
            border-radius: 10px;
            margin-bottom: 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0
