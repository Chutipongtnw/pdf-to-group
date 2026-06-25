import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import unicodedata

# ฟังก์ชันทำความสะอาดข้อความ
def universal_thai_cleaner(text):
    if not text: return "N/A"
    
    for divider in ["จำนวน", "จํานวน", "หน่วย"]:
        if divider in text:
            text = text.split(divider)[0]

    text = unicodedata.normalize('NFKC', text)
    
    text = text.replace('ค', 'ค้น').replace('คว', 'คว้า')
    text = text.replace('ด', 'ด้')
    text = text.replace('สราง', 'สร้าง')
    text = text.replace('รู', 'รู้')
    text = text.replace('หนา', 'หน้า')
    text = text.replace('ตน', 'ต้น')
    text = text.replace('ป่ญญา', 'ปัญญา')
    
    unicode_map = {
        '\uf701': 'ิ', '\uf702': 'ี', '\uf703': 'ึ', '\uf704': 'ื',
        '\uf705': '่', '\uf706': '้', '\uf70e': '์', '\uf710': '่',
        '\uf711': '้', '\uf714': '์', '\uf71a': '์', '\uf709': '',
        '\uf712': 'เ', '\uf713': 'เ',
        'อ': 'อ่าน', 'ข': 'ข้อ', 'ค': 'ค้น', 'ต': 'ต่อ', 'นํ': 'นำ', 'ผ': 'แผ่น'
    }
    for char, corrected in unicode_map.items():
        text = text.replace(char, corrected)

    text = re.sub(r'[\u0000-\u001f\u007f-\u009f\uf000-\uf0ff\u200b\u00a0]', '', text)
    text = "".join(text.split())

    text = text.replace('ศลิป', 'ศิลป์') 
    text = text.replace('ฟิสกิ', 'ฟิสิก') 
    text = text.replace('ต่ออ', 'ต่อ')
    text = text.replace('ค้นน', 'ค้น')
    text = text.replace('ข้ออ', 'ข้อ')
    text = text.replace('แผ่นน', 'แผ่น')
    text = text.replace('นำา', 'นำ')
    text = text.replace('ฟ่ง', 'ฟัง')
    text = text.replace('อ่านาน', 'อ่าน') 
    
    for _ in range(2):
        text = text.replace('เเ', 'เ')
        text = text.replace('แแ', 'แ')
        text = text.replace('าา', 'า')
    
    if 'พิ่มเติม' in text and 'เพิ่ม' not in text:
        text = text.replace('พิ่มเติม', 'เพิ่มเติม')

    corrections = {
        'ศกึ': 'ศึก',
        'วิทยาศาตร์': 'วิทยาศาสตร์',
        'นาฏศิลป1': 'นาฏศิลป์ 1',
        'นาฏศิลป2': 'นาฏศิลป์ 2',
        'ทัศนศิลป1': 'ทัศนศิลป์ 1',
        'ทัศนศิลป2': 'ทัศนศิลป์ 2',
        'ฟิสิกส': 'ฟิสิกส์',
        'คณิตศาสตร': 'คณิตศาสตร์',
        'ผลติ': 'ผลิต',
        'คาสตร์': 'ศาสตร์',
        'วดีโอ': 'วิดีโอ',
        'เ์': '์'
    }
    for wrong, right in corrections.items():
        text = text.replace(wrong, right)

    text = re.sub(r'([่้๊๋์])\1+', r'\1', text)
    text = re.sub(r'(\d+)$', r' \1', text)

    return text.strip()

# ฟังก์ชันสำหรับเคลียร์ข้อความในตาราง (คงช่องว่างระหว่างชื่อ-นามสกุลไว้)
def clean_row_cell(text):
    if not text: return ""
    return str(text).replace('\n', ' ').strip()

st.set_page_config(page_title="ระบบดึงข้อมูลจากตาราง PDF v57", layout="wide")
st.title("📂 ระบบดึงข้อมูลจากตาราง PDF -> Excel")
st.write("ดึงข้อมูลคอลัมน์: ที่ / เลขประจำตัว / ชื่อ-ชื่อสกุล / ห้อง / หน่วยการเรียนที่เรียน / หน่วยการเรียนที่ได้ / ระดับคะแนนเฉลี่ยเฉพาะกลุ่ม")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF ที่ต้องการแปลง", type="pdf")

if uploaded_file is not None:
    all_data = []
    
    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        total_pages = len(pdf.pages)
        
        for i, page in enumerate(pdf.pages):
            table = page.extract_table()
            if table:
                for row in table:
                    # ตารางต้องมีอย่างน้อย 6 คอลัมน์ ถึงจะเข้าเงื่อนไขทำงานได้
                    if row and len(row) >= 6:
                        # ข้ามแถวที่เป็นหัวตาราง
                        if "เลขประจำตัว" in str(row) or "เลขที่" in str(row):
                            continue
                        
                        # ✨ ดึง 'เลขที่' จากคอลัมน์แรกสุด (Index 0) ของ PDF
                        s_no = str(row[0]).replace('\n', '').strip()
                        # เลขประจำตัว จะอยู่คอลัมน์ที่ 2 (Index 1) เสมอ
                        s_id = str(row[1]).replace('\n', '').strip()
                        
                        # เช็คว่าเป็นรหัสนักเรียนจริงหรือไม่ (ต้องเป็นตัวเลข และมี 4 หลักขึ้นไป)
                        if s_id.isdigit() and len(s_id) >= 4:
                            # ชื่อ-นามสกุล จะอยู่คอลัมน์ที่ 3 (Index 2) เสมอ
                            s_name = clean_row_cell(row[2])
                            
                            # จัดการเคสหน้าแรก: 'ห้อง' กับ 'หน่วยการเรียนที่เรียน' โดนรวบ (เหลือ 6 คอลัมน์)
                            if len(row) == 6:
                                merged_col = str(row[3]).split()
                                if len(merged_col) >= 2:
                                    # แยกตัวเลขที่มีจุดทศนิยมไปเป็นหน่วยกิต อีกตัวให้เป็นห้อง
                                    if '.' in merged_col[0]:
                                        s_credit_reg = merged_col[0]
                                        s_room = merged_col[1]
                                    else:
                                        s_room = merged_col[0]
                                        s_credit_reg = merged_col[1]
                                else:
                                    s_room = clean_row_cell(row[3])
                                    s_credit_reg = clean_row_cell(row[3])
                                    
                                s_credit_earn = clean_row_cell(row[4])
                                s_gpa = clean_row_cell(row[5])
                                
                            # เคสปกติของหน้าอื่นๆ: คอลัมน์แยกกันสมบูรณ์ (7 คอลัมน์)
                            elif len(row) >= 7:
                                s_room = clean_row_cell(row[3])
                                s_credit_reg = clean_row_cell(row[4])
                                s_credit_earn = clean_row_cell(row[5])
                                s_gpa = clean_row_cell(row[6])
                            else:
                                continue
                            
                            # ล้างค่าการเว้นวรรคแปลกๆ ที่อาจติดมากับตัวเลขหน่วยกิตและเกรด
                            s_credit_reg = s_credit_reg.replace(' ', '')
                            s_credit_earn = s_credit_earn.replace(' ', '')
                            s_gpa = s_gpa.replace(' ', '')

                            all_data.append({
                                "ที่": s_no,  # ใช้เลขที่ที่ดึงมาจาก PDF โดยตรง
                                "เลขประจำตัว": s_id,
                                "ชื่อ-ชื่อสกุล": s_name,
                                "ห้อง": s_room,
                                "หน่วยการเรียนที่เรียน": s_credit_reg,
                                "หน่วยการเรียนที่ได้": s_credit_earn,
                                "ระดับคะแนนเฉลี่ยเฉพาะกลุ่ม": s_gpa
                            })
                            
            progress_bar.progress((i + 1) / total_pages)

    if all_data:
        # นำข้อมูลเข้าตาราง Excel
        df = pd.DataFrame(all_data).drop_duplicates().reset_index(drop=True)
        
        st.success(f"ดึงข้อมูลสำเร็จ! พบทั้งหมด {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        # เขียนไฟล์เพื่อเตรียมดาวน์โหลด
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
            
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel", 
            data=output.getvalue(), 
            file_name="student_filtered_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("ไม่พบข้อมูลตารางที่ตรงตามเงื่อนไขในไฟล์ PDF นี้")
