import streamlit as st
import pdfplumber
import os
import tempfile
import zipfile
from pypdf import PdfWriter, PdfReader
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from io import BytesIO

# --- CẤU HÌNH TRANG WEB ---
st.set_page_config(page_title="Tool Tách LO PDF", page_icon="📄")

st.title("📄 Tool Tách File PDF theo LO")
st.markdown("""
**Hướng dẫn:**
1. Upload file PDF "Check Refund Backup".
2. Nhập "Ngày Batch" (nếu cần thay đổi).
3. Bấm **Xử lý** và tải về file ZIP.
""")

# --- HÀM TẠO TRANG 1 (GIỐNG CŨ) ---
def create_page_1(data_row, header_row, temp_filename):
    doc = SimpleDocTemplate(temp_filename, pagesize=landscape(letter))
    elements = []
    styles = getSampleStyleSheet()
    
    # Lấy số LO (cột 7) để làm tiêu đề
    lo_number = data_row[7] if len(data_row) > 7 else "UNKNOWN"
    title = Paragraph(f"<b>LO REFUND DETAIL: {lo_number}</b>", styles['Heading1'])
    elements.append(title)
    elements.append(Spacer(1, 20))

    table_data = [header_row, data_row]
    t = Table(table_data)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
    ]))
    elements.append(t)
    doc.build(elements)

# --- GIAO DIỆN CHÍNH ---
uploaded_file = st.file_uploader("Chọn file PDF gốc", type=["pdf"])

if uploaded_file is not None:
    # Gợi ý ngày batch từ tên file (lấy 6 ký tự đầu)
    default_date = uploaded_file.name[:6] if uploaded_file.name[:6].isdigit() else "112425"
    batch_date = st.text_input("Ngày Batch (để đặt tên file)", value=default_date)

    if st.button("🚀 Xử lý ngay"):
        with st.spinner('Đang tách file... vui lòng chờ'):
            # Tạo thư mục tạm để xử lý
            with tempfile.TemporaryDirectory() as temp_dir:
                # Lưu file upload xuống tạm thời để thư viện đọc
                input_path = os.path.join(temp_dir, "input.pdf")
                with open(input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                # 1. Đọc dữ liệu bảng
                extracted_rows = []
                header = []
                with pdfplumber.open(input_path) as pdf:
                    page1 = pdf.pages[0] # Giả định bảng ở trang 1
                    table = page1.extract_table()
                    if table:
                        start_row_index = 0
                        for i, row in enumerate(table):
                            if row and "Vendor No." in str(row):
                                header = row
                                start_row_index = i + 1
                                break
                        for row in table[start_row_index:]:
                            # Kiểm tra cột LO (index 7) có dữ liệu số không
                            if row and len(row) > 7 and row[7] is not None:
                                clean_lo = row[7].strip().replace('\n', '')
                                if clean_lo.isdigit():
                                    # Cập nhật lại giá trị sạch vào row
                                    row[7] = clean_lo 
                                    extracted_rows.append(row)

                st.write(f"✅ Tìm thấy {len(extracted_rows)} mã LO.")

                # 2. Lấy 2 trang cuối
                reader = PdfReader(input_path)
                total_pages = len(reader.pages)
                if total_pages < 3:
                    st.error("File quá ngắn (< 3 trang), không thể tách.")
                    st.stop()
                
                last_page = reader.pages[total_pages - 1]
                second_last_page = reader.pages[total_pages - 2]

                # 3. Tạo file ZIP trong bộ nhớ
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    
                    # Vòng lặp tạo PDF
                    for row in extracted_rows:
                        lo_id = row[7]
                        pdf_name = f"{batch_date}-{lo_id}.pdf"
                        temp_page1_path = os.path.join(temp_dir, "temp_page1.pdf")
                        
                        try:
                            # Tạo trang 1
                            create_page_1(row, header, temp_page1_path)
                            
                            # Ghép file
                            merger = PdfWriter()
                            merger.add_page(PdfReader(temp_page1_path).pages[0]) # Trang 1 mới
                            merger.add_page(second_last_page) # Trang áp chót cũ
                            merger.add_page(last_page) # Trang cuối cũ

                            # Lưu vào buffer
                            output_pdf_buffer = BytesIO()
                            merger.write(output_pdf_buffer)
                            
                            # Đưa vào file ZIP
                            zip_file.writestr(pdf_name, output_pdf_buffer.getvalue())
                        except Exception as e:
                            st.warning(f"Lỗi khi tạo LO {lo_id}: {e}")

                # 4. Hiển thị nút Download
                st.success("🎉 Xử lý xong!")
                st.download_button(
                    label="📥 Tải xuống tất cả (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"LO_Refunds_{batch_date}.zip",
                    mime="application/zip"
                )