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
st.set_page_config(page_title="Tool Tách LO PDF (V3 Pro)", page_icon="🧩")

st.title("🧩 Tool Tách PDF theo LO (Bản V3 - Grouping)")
st.markdown("""
**Tính năng V3:**
1. **Gộp dòng:** Các dòng có cùng mã LO sẽ được gộp chung vào 1 file.
2. **Trang sau:** Giữ nguyên tất cả các trang từ trang 2 trở đi của file gốc.
3. **Tự động:** Nhận diện cột LO thông minh hơn.
""")

# --- HÀM TẠO TRANG 1 (VẼ LẠI BẢNG VỚI NHIỀU DÒNG) ---
def create_page_1_group(data_rows, header_row, temp_filename, lo_number):
    """
    Tạo trang 1 mới chứa Header và DANH SÁCH các dòng dữ liệu của LO đó
    """
    doc = SimpleDocTemplate(temp_filename, pagesize=landscape(letter))
    elements = []
    styles = getSampleStyleSheet()
    
    # Tiêu đề
    title = Paragraph(f"<b>LO REFUND DETAIL: {lo_number}</b>", styles['Heading1'])
    elements.append(title)
    elements.append(Spacer(1, 20))

    # Chuẩn bị dữ liệu bảng: Header + Các dòng dữ liệu
    # Làm sạch dữ liệu để tránh lỗi hiển thị
    clean_header = [str(h).replace('\n', ' ') if h else "" for h in header_row]
    
    table_data = [clean_header] # Dòng đầu tiên là header
    
    for row in data_rows:
        # Làm sạch từng ô trong dòng
        clean_row = [str(cell).replace('\n', ' ') if cell else "" for cell in row]
        table_data.append(clean_row)

    # Tạo bảng
    # Tính toán độ rộng cột tự động (đơn giản hóa) hoặc để tự động
    t = Table(table_data)
    
    # Định dạng bảng (Style đẹp mắt)
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.2, 0.2, 0.2)), # Màu xám đậm cho Header
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke), # Chữ trắng cho Header
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey), # Kẻ khung mỏng
        ('FONTSIZE', (0, 1), (-1, -1), 8), # Cỡ chữ dữ liệu nhỏ hơn xíu
        ('LEFTPADDING', (0,0), (-1,-1), 3),
        ('RIGHTPADDING', (0,0), (-1,-1), 3),
    ])
    
    # Tô màu xen kẽ các dòng dữ liệu cho dễ nhìn
    for i in range(1, len(table_data)):
        if i % 2 == 0:
            bg_color = colors.whitesmoke
        else:
            bg_color = colors.Color(0.95, 0.95, 0.9, 1) # Màu beige nhạt
        style.add('BACKGROUND', (0, i), (-1, i), bg_color)

    t.setStyle(style)
    elements.append(t)
    doc.build(elements)

# --- GIAO DIỆN CHÍNH ---
uploaded_file = st.file_uploader("Chọn file PDF gốc", type=["pdf"])

if uploaded_file is not None:
    # Gợi ý ngày batch từ tên file
    default_date = "MMDDYY"
    if len(uploaded_file.name) >= 6 and uploaded_file.name[:6].isdigit():
        default_date = uploaded_file.name[:6]
        
    batch_date = st.text_input("Ngày Batch (để đặt tên file)", value=default_date)

    if st.button("🚀 Xử lý ngay"):
        with st.spinner('Đang phân tích bảng dữ liệu...'):
            with tempfile.TemporaryDirectory() as temp_dir:
                # Lưu file tạm để xử lý
                input_path = os.path.join(temp_dir, "input.pdf")
                with open(input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                header = []
                lo_col_index = -1
                lo_groups = {} # Dictionary để gom nhóm: {'016': [row1, row2], ...}

                # --- GIAI ĐOẠN 1: ĐỌC DỮ LIỆU TỪ TRANG 1 ---
                with pdfplumber.open(input_path) as pdf:
                    page1 = pdf.pages[0]
                    # Thử extract table với cài đặt mặc định
                    table = page1.extract_table()
                    
                    if not table:
                        # Nếu thất bại, thử chế độ khác
                        table = page1.extract_table(table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"})

                    if table:
                        # 1. Tìm dòng Header và vị trí cột LO
                        header_row_idx = -1
                        for i, row in enumerate(table):
                            # Chuyển row thành text để tìm kiếm từ khóa
                            row_text_list = [str(c).strip().upper() for c in row if c]
                            
                            # Tiêu chí nhận diện Header: Có chữ "LO" và ("AMOUNT" hoặc "NAME" hoặc "SALE")
                            if "LO" in row_text_list:
                                header = row
                                header_row_idx = i
                                
                                # Tìm index của cột LO
                                for idx, col_name in enumerate(row):
                                    if col_name and "LO" == str(col_name).strip().upper():
                                        lo_col_index = idx
                                        break
                                break
                        
                        if lo_col_index == -1:
                            st.error("❌ Không tìm thấy cột 'LO' trong bảng. Hãy kiểm tra file PDF.")
                            st.write("Dữ liệu đọc được:", table[:5])
                            st.stop()
                        
                        # 2. Quét dữ liệu bên dưới Header
                        count_found = 0
                        for row in table[header_row_idx + 1:]:
                            # Dòng phải đủ dài và cột LO không được rỗng
                            if row and len(row) > lo_col_index:
                                raw_lo = row[lo_col_index]
                                if raw_lo:
                                    # Làm sạch mã LO: xóa xuống dòng, khoảng trắng
                                    clean_lo = str(raw_lo).strip().replace('\n', '')
                                    
                                    # Logic nhận diện LO: Là số (016, 235...) hoặc dạng chuỗi đặc biệt nếu cần
                                    # Ở đây ta lấy tất cả nếu nó trông giống mã số
                                    if clean_lo.isdigit(): 
                                        # Cập nhật lại giá trị sạch vào row để in ra đẹp
                                        row[lo_col_index] = clean_lo
                                        
                                        # Thêm vào nhóm
                                        if clean_lo not in lo_groups:
                                            lo_groups[clean_lo] = []
                                        lo_groups[clean_lo].append(row)
                                        count_found += 1
                        
                        if count_found == 0:
                            st.warning("⚠️ Tìm thấy cột LO nhưng không có dòng dữ liệu nào bên dưới chứa số.")
                            st.stop()
                            
                    else:
                        st.error("❌ Không đọc được bảng từ trang 1 PDF.")
                        st.stop()

                st.success(f"✅ Đã tìm thấy {len(lo_groups)} mã LO (đã gộp các dòng trùng).")

                # --- GIAI ĐOẠN 2: TẠO FILE PDF KẾT QUẢ ---
                reader = PdfReader(input_path)
                total_pages = len(reader.pages)
                
                # Tạo file ZIP trong bộ nhớ
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    
                    # Duyệt qua từng nhóm LO để tạo file
                    for lo_id, rows in lo_groups.items():
                        pdf_name = f"{batch_date}-{lo_id}.pdf"
                        temp_page1_path = os.path.join(temp_dir, "temp_page1.pdf")
                        
                        try:
                            # A. Tạo trang 1 mới (chứa bảng đã gộp các rows)
                            create_page_1_group(rows, header, temp_page1_path, lo_id)
                            
                            # B. Ghép file
                            merger = PdfWriter()
                            
                            # 1. Thêm trang 1 vừa tạo
                            merger.add_page(PdfReader(temp_page1_path).pages[0])
                            
                            # 2. Thêm TẤT CẢ các trang còn lại từ file gốc (Từ trang 2 -> Hết)
                            # Lưu ý: Index trong PyPDF bắt đầu từ 0. Trang 2 là index 1.
                            if total_pages > 1:
                                for i in range(1, total_pages):
                                    merger.add_page(reader.pages[i])

                            # C. Lưu vào ZIP
                            output_pdf_buffer = BytesIO()
                            merger.write(output_pdf_buffer)
                            zip_file.writestr(pdf_name, output_pdf_buffer.getvalue())
                            
                        except Exception as e:
                            st.warning(f"Lỗi khi tạo file LO {lo_id}: {e}")

                # Nút tải về
                st.download_button(
                    label=f"📥 Tải xuống {len(lo_groups)} file (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"LO_Refunds_{batch_date}.zip",
                    mime="application/zip"
                )
