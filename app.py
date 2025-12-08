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
st.set_page_config(page_title="Tool Tách LO PDF (V2)", page_icon="🛠️")

st.title("🛠️ Tool Tách File PDF theo LO (Bản V2)")
st.info("Phiên bản này có tính năng 'Dò tìm thông minh' để sửa lỗi không thấy mã LO.")

# --- HÀM TẠO TRANG 1 ---
def create_page_1(data_row, header_row, temp_filename, lo_index):
    doc = SimpleDocTemplate(temp_filename, pagesize=landscape(letter))
    elements = []
    styles = getSampleStyleSheet()
    
    # Lấy số LO từ vị trí cột đã tìm thấy
    lo_number = data_row[lo_index] if len(data_row) > lo_index else "UNKNOWN"
    title = Paragraph(f"<b>LO REFUND DETAIL: {lo_number}</b>", styles['Heading1'])
    elements.append(title)
    elements.append(Spacer(1, 20))

    # Chỉ lấy các cột có dữ liệu để bảng đẹp hơn (tránh cột None)
    clean_header = [h if h else "" for h in header_row]
    clean_data = [d if d else "" for d in data_row]

    table_data = [clean_header, clean_data]
    t = Table(table_data)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTSIZE', (0, 0), (-1, -1), 7), # Giảm font xíu để vừa bảng
    ]))
    elements.append(t)
    doc.build(elements)

# --- GIAO DIỆN CHÍNH ---
uploaded_file = st.file_uploader("Chọn file PDF gốc", type=["pdf"])

if uploaded_file is not None:
    # Gợi ý ngày batch
    default_date = uploaded_file.name[:6] if uploaded_file.name[:6].isdigit() else "112425"
    batch_date = st.text_input("Ngày Batch", value=default_date)

    if st.button("🚀 Xử lý ngay"):
        with st.spinner('Đang phân tích bảng dữ liệu...'):
            with tempfile.TemporaryDirectory() as temp_dir:
                input_path = os.path.join(temp_dir, "input.pdf")
                with open(input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                extracted_rows = []
                header = []
                lo_col_index = -1 # Chưa tìm thấy

                # 1. Đọc và gỡ lỗi (Debug) dữ liệu bảng
                with pdfplumber.open(input_path) as pdf:
                    page1 = pdf.pages[0]
                    # Thử chế độ snap=True để bắt bảng tốt hơn
                    table = page1.extract_table(table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"})
                    
                    if not table:
                        # Thử lại với chế độ mặc định nếu chế độ text fail
                        table = page1.extract_table()

                    if table:
                        # --- LOGIC DÒ TÌM HEADER THÔNG MINH ---
                        header_row_idx = -1
                        st.write("🔍 **Đang kiểm tra cấu trúc bảng...**")
                        
                        for i, row in enumerate(table):
                            # Làm sạch row để tìm kiếm
                            row_str = [str(c).strip() for c in row if c]
                            
                            # Tìm dòng chứa chữ "LO" (Đây là dấu hiệu nhận biết header)
                            if "LO" in row_str:
                                header = row
                                header_row_idx = i
                                
                                # Tìm vị trí cột LO nằm ở đâu
                                for idx, col_name in enumerate(row):
                                    if col_name and "LO" == col_name.strip():
                                        lo_col_index = idx
                                        break
                                
                                st.success(f"✅ Đã tìm thấy Header ở dòng {i+1}. Cột LO nằm ở vị trí số {lo_col_index+1}")
                                break
                        
                        if lo_col_index == -1:
                            st.error("❌ Không tìm thấy cột nào tên là 'LO' trong bảng. Vui lòng kiểm tra lại file PDF.")
                            st.write("Dữ liệu 5 dòng đầu tiên đọc được là:")
                            st.write(table[:5]) # In ra để debug
                            st.stop()

                        # --- LẤY DỮ LIỆU ---
                        for row in table[header_row_idx + 1:]:
                            # Phải có đủ số cột và cột LO không được trống
                            if row and len(row) > lo_col_index:
                                raw_lo = row[lo_col_index]
                                if raw_lo:
                                    clean_lo = str(raw_lo).strip().replace('\n', '')
                                    # Chấp nhận nếu là số (ví dụ '016')
                                    if clean_lo.isdigit():
                                        row[lo_col_index] = clean_lo
                                        extracted_rows.append(row)
                    else:
                        st.error("❌ Không đọc được bảng nào từ trang 1 PDF.")
                        st.stop()

                st.write(f"📊 **Kết quả:** Tìm thấy {len(extracted_rows)} dòng dữ liệu hợp lệ.")

                if not extracted_rows:
                    st.warning("⚠️ Không có dòng dữ liệu nào bên dưới Header có chứa mã LO là số.")
                    st.stop()

                # 2. Xử lý tách file
                reader = PdfReader(input_path)
                if len(reader.pages) < 3:
                    st.error("File quá ngắn (< 3 trang).")
                    st.stop()
                
                last_page = reader.pages[-1]
                second_last_page = reader.pages[-2]

                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for row in extracted_rows:
                        lo_id = row[lo_col_index]
                        pdf_name = f"{batch_date}-{lo_id}.pdf"
                        temp_page1_path = os.path.join(temp_dir, "temp_page1.pdf")
                        
                        try:
                            # Truyền thêm lo_col_index vào hàm tạo trang
                            create_page_1(row, header, temp_page1_path, lo_col_index)
                            
                            merger = PdfWriter()
                            merger.add_page(PdfReader(temp_page1_path).pages[0])
                            merger.add_page(second_last_page)
                            merger.add_page(last_page)

                            output_pdf_buffer = BytesIO()
                            merger.write(output_pdf_buffer)
                            zip_file.writestr(pdf_name, output_pdf_buffer.getvalue())
                        except Exception as e:
                            st.warning(f"Lỗi tạo file {lo_id}: {e}")

                st.download_button(
                    label="📥 Tải xuống tất cả (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"LO_Refunds_{batch_date}.zip",
                    mime="application/zip"
                )
