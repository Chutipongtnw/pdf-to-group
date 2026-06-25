import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import unicodedata

def universal_thai_cleaner(text):
    if not text: return "N/A"
    
    # 1. ตัดส่วน "จำนวน" ออกทันที
    for divider in ["จำนวน", "จํานวน", "หน่วย"]:
        if divider in text:
            text = text.split(divider)[0]

    # 2. Normalization ขั้นสูงสุด
    text = unicodedata.normalize('NFKC', text)
    
    # 3. Unicode Mapping และซ่อมคำเฉพาะหน้า
    text = text.replace('ค', 'ค้น').replace('คว', 'คว้า')
    text = text.replace('ด', 'ด้')
    
    # --- ซ่อมคำเฉพาะ 5 วิชาโดยตรง ไม่กระทบส่วนอื่น ---
    text = text.replace('สราง', 'สร้าง')
    text = text.replace('รู', 'รู้')
    text = text.replace('หนา', 'หน้า')
    text = text.replace('ตน', 'ต้น')
    text = text.replace('ป่ญญา', 'ปัญญา')
    # --------------------------------------------------------
    
    unicode_map = {
        '\uf701': 'ิ', '\uf702': 'ี', '\uf703': 'ึ', '\uf704': 'ื',
        '\uf705': '่', '\uf706': '้', '\uf70e': '์', '\uf710': '่',
        '\uf711': '้', '\uf714': '์', '\uf71a': '์', '\uf709': '',
        '\uf712': 'เ', '\uf713': 'เ',
        'อ': 'อ่าน', 'ข': 'ข้อ', 'ค': 'ค้น', 'ต': 'ต่อ', 'นํ': 'นำ', 'ผ': 'แผ่น'
    }
    for char, corrected in unicode_map.items():
        text = text.replace(char, corrected)

    # 4. ล้างรหัส Unicode ขยะและช่องว่างทั้งหมด
    text = re.sub(r'[\u0000-\u001f\u007f-\u009f\uf000-\uf0ff\u200b\u00a0]', '', text)
    text = "".join(text.split())

    # 5. แก้ไขพยัญชนะเบิ้ลและสระกระโดด
    text = text.replace('ศลิป', 'ศิลป์') 
    text = text.replace('ฟิสกิ', 'ฟิสิก') 
    text = text.replace('ต่ออ', 'ต่อ')
    text = text.replace('ค้นน', 'ค้น')
    text = text.replace('ข้ออ', 'ข้อ')
    text = text.replace('แผ่นน', 'แผ่น')
    text = text.replace('นำา', 'นำ')
    text = text.replace('ฟ่ง', 'ฟัง')
    text = text.replace('อ่านาน', 'อ่าน') 
    
    # 6. ยุบสระที่เบิ้ล (เเ, แแ, าา)
    for _ in range(2):
        text = text.replace('เเ', 'เ')
        text = text.replace('แแ', 'แ')
        text = text.replace('าา', 'า')
    
    # 7. คลังซ่อมคำมาตรฐาน
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

    # 8. ยุบวรรณยุกต์ซ้ำ
    text = re.sub(r'([่้๊๋์])\1+', r'\1', text)
    
    # 9. คืนค่าช่องว่าง 1 เคาะ หน้าตัวเลขท้ายชื่อวิชา
    text = re.sub(r'(\d+)$', r' \1', text)

    return text.strip()

# ฟังก์ชันสำหรับเคลียร์ข้อความในตารางเบื้องต้น (คงเหลือช่องว่างไว้สำหรับ ชื่อ-นามสกุล)
def clean_row_cell(text):
    if not text: return ""
    text = str(text).replace('\n', ' ').strip()
    return text

st.set_page_config(page_title="ระบบดึงข้อมูลจากตาราง PDF v55", layout="wide")
st.title("📂 ระบบดึงข้อมูลจากตาราง PDF -> Excel")
st.write("ดึงข้อมูลคอลัมน์: เลขประจำตัว / ชื่อ-ชื่อสกุล / ห้อง / หน่วยการเรียนที่เรียน / หน่วยการเรียนที่ได้ / ระดับคะแนนเฉลี่ยเฉพาะกลุ่ม")

uploaded_file = st.file_uploader("เลือกไฟล์ PDF ที่ต้องการแปลง", type="pdf")

if uploaded_file is not None:
    all_data = []
    
    # =========================================================================
    # 🛠️ [ส่วนสำคัญ] ตั้งค่าตำแหน่งคอลัมน์ที่อยู่ใน PDF จริง 
    # เริ่มนับจาก 0 (0 คือคอลัมน์แรกสุดซ้ายมือ, 1 คือคอลัมน์ถัดไป...)
    # หากตำแหน่งใน PDF เปลี่ยนไป ให้มาแก้ตัวเลขดัชนี (Index) ตรงนี้ได้เลยครับ
    # =========================================================================
    IDX_STUDENT_ID = 0   # ตำแหน่งของ "เลขประจำตัว"
    IDX_NAME = 1         # ตำแหน่งของ "ชื่อ-ชื่อสกุล"
    IDX_ROOM = 2         # ตำแหน่งของ "ห้อง"
    IDX_CREDIT_REG = 3   # ตำแหน่งของ "หน่วยการเรียนที่เรียน"
    IDX_CREDIT_EARN = 4  # ตำแหน่งของ "หน่วยการเรียนที่ได้"
    IDX_GPA_GROUP = 5    # ตำแหน่งของ "ระดับคะแนนเฉลี่ยเฉพาะกลุ่ม"
    # =========================================================================

    with pdfplumber.open(uploaded_file) as pdf:
        progress_bar = st.progress(0)
        total_pages = len(pdf.pages)
        
        for i, page in enumerate(pdf.pages):
            table = page.extract_table()
            if table:
                for row in table:
                    # ตรวจสอบว่าแถวมีข้อมูลและมีจำนวนคอลัมน์เพียงพอกับที่ตั้งค่าไว้หรือไม่
                    max_idx = max(IDX_STUDENT_ID, IDX_NAME, IDX_ROOM, IDX_CREDIT_REG, IDX_CREDIT_EARN, IDX_GPA_GROUP)
                    if row and len(row) > max_idx:
                        
                        # ดึงข้อมูลตามตำแหน่ง Index ที่ตั้งค่าไว้
                        s_id = clean_row_cell(row[IDX_STUDENT_ID])
                        s_name = clean_row_cell(row[IDX_NAME])
                        s_room = clean_row_cell(row[IDX_ROOM])
                        s_credit_reg = clean_row_cell(row[IDX_CREDIT_REG])
                        s_credit_earn = clean_row_cell(row[IDX_CREDIT_EARN])
                        s_gpa_group = clean_row_cell(row[IDX_GPA_GROUP])
                        
                        # กรองเอาเฉพาะแถวที่เป็นข้อมูลนักเรียนจริง (หลีกเลี่ยงแถวหัวข้อตาราง หรือแถวว่าง)
                        # ในที่นี้ตรวจสอบว่า "เลขประจำตัว" ต้องเป็นตัวเลข และไม่เป็นค่าว่าง
                        if s_id.isdigit() and len(s_id) >= 4:
                            all_data.append({
                                "เลขประจำตัว": s_id,
                                "ชื่อ-ชื่อสกุล": s_name,
                                "ห้อง": s_room,
                                "หน่วยการเรียนที่เรียน": s_credit_reg,
                                "หน่วยการเรียนที่ได้": s_credit_earn,
                                "ระดับคะแนนเฉลี่ยเฉพาะกลุ่ม": s_gpa_group
                            })
                            
            progress_bar.progress((i + 1) / total_pages)

    if all_data:
        # แปลงเป็น DataFrame และลบรายการที่ซ้ำออก (ถ้ามี)
        df = pd.DataFrame(all_data).drop_duplicates().reset_index(drop=True)
        
        # ✨ เพิ่มคอลัมน์ "ที่" เรียงลำดับลงไปเรื่อยๆ เริ่มจาก 1 ไว้ข้างหน้าสุด
        df.insert(0, "ที่", range(1, len(df) + 1))
        
        st.success(f"ดึงข้อมูลสำเร็จ! พบทั้งหมด {len(df)} รายการ")
        st.dataframe(df, use_container_width=True)
        
        # เขียนข้อมูลลงไฟล์ Excel
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
        st.warning("ไม่พบข้อมูลตารางที่ตรงตามเงื่อนไขในไฟล์ PDF นี้ กรุณาตรวจสอบตำแหน่งคอลัมน์ (Index) ในโค้ด")
