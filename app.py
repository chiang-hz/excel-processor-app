# app.py

import streamlit as st
import pandas as pd
from fpdf import FPDF
from io import BytesIO
import os

# --- 1. 核心處理函式 (跨平台版本) ---
# 使用 pandas 和 fpdf2 來處理檔案，不依賴 pywin32

class PDF(FPDF):
    """
    繼承 FPDF 類別以自訂頁首和頁尾。
    """
    def __init__(self, orientation='P', unit='mm', format='A4', footer_text=""):
        super().__init__(orientation, unit, format)
        self.footer_text = footer_text
        self.set_auto_page_break(auto=True, margin=15)

    def footer(self):
        # 設定頁尾位置在頁面底部 1.5 公分處
        self.set_y(-15)
        # 設定字體
        self.set_font('NotoSans', '', 8)
        # 建立頁碼文字
        page_num_text = f"Page {self.page_no()}/{{nb}}"
        
        # 使用者自訂的文字與頁碼結合
        final_footer_text = self.footer_text.replace('&P', str(self.page_no())).replace('&N', '{nb}')
        
        # 輸出頁尾
        self.cell(0, 10, final_footer_text, 0, 0, 'C')

def process_excel_to_pdf_cross_platform(uploaded_file, options: dict) -> bytes:
    """
    使用 pandas 讀取 Excel 數據，使用 fpdf2 產生 PDF。

    Args:
        uploaded_file: Streamlit 的 UploadedFile 物件。
        options (dict): 包含排版設定的字典。

    Returns:
        bytes: PDF 檔案的二進位內容。
    """
    try:
        # --- 讀取所有 Excel 工作表 ---
        # sheet_name=None 會讀取所有工作表到一個字典中
        xls = pd.ExcelFile(uploaded_file)
        sheets_data = {sheet_name: xls.parse(sheet_name) for sheet_name in xls.sheet_names}

        # --- 初始化 PDF 物件 ---
        # fpdf2 使用 'mm' 作為單位，我們將 cm 轉換為 mm
        orientation_map = {'直向': 'P', '橫向': 'L'}
        pdf = PDF(
            orientation=orientation_map.get(options['頁面方向'], 'P'),
            unit='mm',
            format=options['紙張大小'],
            footer_text=options['頁尾內容']
        )
        
        # --- 字體設定 (支援中文) ---
        # 為了在 Streamlit Cloud 上運行，我們需要提供字體檔
        # 請下載 'NotoSansTC-Regular.ttf' 並和 app.py 放在同一個資料夾
        font_path = 'NotoSansTC-Regular.ttf'
        if not os.path.exists(font_path):
            st.error(f"錯誤：找不到字體檔案 '{font_path}'。請下載並將其與 app.py 放在一起。")
            st.error("您可以從 Google Fonts 下載：https://fonts.google.com/noto/specimen/Noto+Sans+TC")
            return None
            
        pdf.add_font('NotoSans', '', font_path, uni=True)
        pdf.set_font('NotoSans', '', 10)

        # 設定頁邊距 (cm to mm)
        pdf.set_top_margin(options['上邊距'] * 10)
        pdf.set_bottom_margin(options['下邊距'] * 10)
        pdf.set_left_margin(options['左邊距'] * 10)
        pdf.set_right_margin(options['右邊距'] * 10)

        # 遍歷每個工作表並將其加入 PDF
        for sheet_name, df in sheets_data.items():
            pdf.add_page()
            
            # 加入工作表標題
            pdf.set_font('NotoSans', '', 14)
            pdf.cell(0, 10, sheet_name, 0, 1, 'L')
            pdf.set_font('NotoSans', '', 10)
            
            # 將 DataFrame 的標頭轉換為字串
            headers = [str(col) for col in df.columns]
            
            # --- 建立表格 ---
            # 設定標頭樣式
            pdf.set_fill_color(230, 230, 230)
            pdf.set_font('NotoSans', '', 10)
            
            # 計算欄寬 (簡易平均分配)
            # 可用的頁面寬度 = 紙張寬度 - 左邊距 - 右邊距
            page_width = pdf.w - pdf.l_margin - pdf.r_margin
            col_width = page_width / len(headers) if headers else page_width

            # 繪製標頭
            for header in headers:
                pdf.cell(col_width, 8, header, 1, 0, 'C', fill=True)
            pdf.ln()

            # 繪製內容
            pdf.set_font('NotoSans', '', 9)
            for index, row in df.iterrows():
                # 將每一行的內容轉換為字串
                str_row = [str(item) for item in row]
                for item in str_row:
                    pdf.multi_cell(col_width, 6, item, border=1, align='L')
                # multi_cell 後需要手動移動到下一行開始的位置
                current_x = pdf.l_margin
                current_y = pdf.get_y() - 6 # 回到儲存格的頂部
                pdf.set_xy(current_x + col_width * (len(str_row)), current_y)
                pdf.ln()

        # 使用別名來計算總頁數
        pdf.alias_nb_pages()
        
        # 輸出 PDF 到一個位元組流
        return pdf.output(dest='S').encode('latin-1')

    except Exception as e:
        st.error(f"處理檔案時發生錯誤: {e}")
        return None

# --- 2. Streamlit 應用程式介面 (UI) ---

st.set_page_config(page_title="Excel to PDF 跨平台轉換器", layout="centered")

st.title("📄 Excel to PDF 跨平台轉換器")
st.info("此版本使用跨平台函式庫，可在任何環境（包括 Streamlit Cloud）執行。它會將數據轉換為標準化的 PDF 表格。")

# --- 3. 側邊欄設定介面 ---
with st.sidebar:
    st.header("⚙️ 排版設定")
    
    uploaded_file = st.file_uploader(
        "上傳 Excel 檔案", 
        type=['xlsx', 'xls'],
        help="支援 .xlsx 和 .xls 格式的檔案。"
    )

    with st.form(key="settings_form"):
        st.subheader("頁面配置")
        paper_size = st.selectbox("紙張大小", options=['A4', 'A3', 'Letter'], index=0)
        page_orientation = st.radio("頁面方向", options=['直向', '橫向'], index=0, horizontal=True)

        st.subheader("頁面邊距 (公分)")
        col1, col2 = st.columns(2)
        with col1:
            top_margin = st.number_input("上邊距", value=1.9, min_value=0.5, step=0.1, format="%.1f")
            bottom_margin = st.number_input("下邊距", value=1.5, min_value=0.5, step=0.1, format="%.1f")
        with col2:
            left_margin = st.number_input("左邊距", value=1.2, min_value=0.5, step=0.1, format="%.1f")
            right_margin = st.number_input("右邊距", value=1.2, min_value=0.5, step=0.1, format="%.1f")
        
        st.subheader("頁尾設定")
        # 替換 &C, &P, &N 為 fpdf2 能理解的格式
        footer_text = st.text_input("頁尾內容", value="第 &P 頁 / 共 &N 頁", help="使用 &P 代表頁碼, &N 代表總頁數。")

        # "縮放模式" 選項已被移除，因為無法在跨平台方案中可靠地實現
        st.info("注意：'縮放模式' 為 Excel 獨有功能，在此版本中不適用。PDF 將根據內容自動排版。")
        
        submit_button = st.form_submit_button(label="🚀 產生 PDF")

# --- 4. 主程式邏輯 ---

if submit_button:
    if uploaded_file is not None:
        with st.spinner("正在轉換為 PDF，請稍候..."):
            # 收集所有設定
            format_options = {
                '紙張大小': paper_size,
                '頁面方向': page_orientation,
                '上邊距': top_margin,
                '下邊距': bottom_margin,
                '左邊距': left_margin,
                '右邊距': right_margin,
                '頁尾內容': footer_text,
            }

            # 呼叫新的跨平台處理函式
            pdf_bytes = process_excel_to_pdf_cross_platform(uploaded_file, format_options)

            if pdf_bytes:
                st.success("PDF 產生成功！")
                
                # 準備下載檔名
                file_name = f"{os.path.splitext(uploaded_file.name)[0]}_converted.pdf"
                
                st.download_button(
                    label="📥 下載 PDF",
                    data=pdf_bytes,
                    file_name=file_name,
                    mime="application/pdf"
                )
    else:
        st.warning("請先上傳一個 Excel 檔案！")