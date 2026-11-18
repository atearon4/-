# -*- coding: utf-8 -*-
"""
CIC LOGIS BOOKING CONSOLE (모바일 웹/PC 겸용 부킹 리스트 관리 시스템)
"""

import os
import re
import csv
import io
import datetime
import urllib.request
import urllib.error
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# PDF 생성 라이브러리 연동
try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib import colors
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

st.set_page_config(page_title="CIC LOGIS 부킹 콘솔", page_icon="📦", layout="wide")

# 스타일 설정
THEME_COLOR = "#0F172A"

# 색상 매핑
CLIENT_COLOR_MAP = {
    "TLC": "#D0E1ED",
    "알프스": "#D4EDDA",
    "서하": "#F8D7DA"
}

FORWARDER_COLOR_MAP = {
    "우정": "#D1ECF1",
    "코스모": "#FFE8D6",
    "국양": "#D6C1B3",
    "트랜스올": "#D4EDDA",
    "바롬": "#E2D9F3"
}

COMPANY_INFO = {
    "NAME_EN": "CIC LOGIS CO.,LTD.",
    "NAME_KR": "(주)씨아이씨로지스틱",
    "ADDR_EN": "D4-block, 95, Jayumuyeok-ro, Jung-gu, Incheon, Korea",
    "ADDR_KR": "인천광역시 중구 자유무역로 95 D4블록 301호",
    "TEL": "+82-32-744-4221~6",
    "FAX": "+82-70-4015-5833",
    "EMAIL": "ciclogis@ciclogis.co.kr"
}

def get_pdf_korean_font():
    font_paths = [
        "C:/Windows/Fonts/malgun.ttf",
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "/System/Library/Fonts/AppleSDGothicNeo.ttc"
    ]
    for p in font_paths:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont("KoreanFont", p))
                return "KoreanFont"
            except Exception:
                continue
    return "Helvetica"

def draw_cic_letterhead_background(canvas, doc):
    canvas.saveState()
    w, h = doc.pagesize

    foot_x = 20
    foot_y = 16
    canvas.setFont("Helvetica-Bold", 7.5)
    canvas.setFillColor(colors.HexColor("#2563EB"))
    canvas.drawString(foot_x, foot_y + 18, COMPANY_INFO["NAME_EN"])

    kor_font = get_pdf_korean_font()
    canvas.setFont(kor_font, 6.5)
    canvas.setFillColor(colors.HexColor("#F43F5E"))
    canvas.drawString(foot_x, foot_y + 10, f"{COMPANY_INFO['ADDR_EN']}  /  {COMPANY_INFO['ADDR_KR']}")
    canvas.drawString(foot_x, foot_y + 2, f"Tel: {COMPANY_INFO['TEL']}   |   Fax: {COMPANY_INFO['FAX']}   |   Email: {COMPANY_INFO['EMAIL']}")

    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(colors.HexColor("#94A3B8"))
    canvas.drawRightString(w - 20, foot_y + 2, f"Page {canvas.getPageNumber()} | CIC LOGIS BOOKING CONSOLE")

    canvas.restoreState()

st.markdown(f"""
    <div style="background-color:{THEME_COLOR}; padding:15px; border-radius:8px; color:white; margin-bottom:20px;">
        <h3 style="margin:0;">📦 CIC LOGIS BOOKING CONSOLE (모바일/PC 겸용)</h3>
        <p style="margin:5px 0 0 0; font-size:13px; color:#94A3B8;">부킹 리스트 전용 관리 및 파일 저장 시스템</p>
    </div>
""", unsafe_allow_html=True)

# 세션 상태 초기화
if "booking_records" not in st.session_state:
    st.session_state.booking_records = []

with st.sidebar:
    st.header("⚙️ 데이터 불러오기")
    fetch_mode = st.radio("방식 선택", ["구글 시트 연동", "엑셀 파일 업로드"])
    
    sheet_url = ""
    sheet_name = "시트1"
    uploaded_file = None

    if fetch_mode == "구글 시트 연동":
        sheet_url = st.text_input("구글 시트 주소 (URL)")
        sheet_name = st.text_input("시트 탭 이름", value="시트1")
        fetch_btn = st.button("🌐 구글시트 가져오기", use_container_width=True)
    else:
        uploaded_file = st.file_uploader("엑셀 파일 선택 (.xlsx)", type=["xlsx", "xls"])
        fetch_btn = st.button("📁 엑셀 파일 열기", use_container_width=True)

    st.divider()
    st.header("🎨