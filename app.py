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
st.set_page_config(page_title="Tool Tách LO PDF (V3 - Grouping)", page_icon="🧩")

st.title("🧩 Tool Tách PDF theo LO (Bản V3)")
st.markdown("""
**Cập nhật mới:**
1. **Gộp dòng:** Các dòng có cùng mã LO sẽ nằm chung trong 1 file.
2. **Cấu trúc:** Trang 1 (Dữ liệu LO) + Các trang còn lại của file gốc (Giữ nguyên).
""")

# --- HÀM TẠO TRANG 1 (VẼ LẠI BẢNG) ---
def create_page_1(data_rows, header_row, temp_filename, lo_number):
    """
    Tạo trang 1 mới chứa Header và danh sách các dòng dữ liệu của LO đó
    """
    doc = SimpleDocTemplate(temp_filename, pagesize=landscape(letter))
    elements = []
    styles = getSampleStyleSheet()
    
    # Tiêu đề
    title = Paragraph(f"<b>LO REFUND DETAIL: {lo_number}</b>", styles['Heading1'])
    elements.append(title)
    elements.append(Spacer(1, 20))

    # Chuẩn bị dữ liệu bảng: Header + Các dòng dữ liệu
    # Làm sạch dữ liệu để tránh lỗi hiển thị None
    clean_header = [str(h) if h else "" for h in header_row]
    table_data = [clean_header] # Dòng đầu là header
    
    for row in data_rows:
        clean_row = [str(d) if d else "" for d in row]
        table_data.append(clean_row)

    # Tạo bảng
    t = Table(table_data)
    
    # Định dạng bảng (Style)
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey), # Màu nền Header
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke), # Màu chữ Header
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black), # Kẻ khung
        ('FONTSIZE', (0, 0), (-1, -1), 8), # Cỡ chữ
    ])
    
    # Tô màu xen kẽ cho các dòng dữ liệu
    for i in range(1, len(table_data)):
        if i % 2 == 0:
            bg_color = colors.whitesmoke
        else:
            bg_color = colors.beige
        style.add('BACKGROUND', (0, i), (-1, i), bg_color)

    t.setStyle(style)
    elements.append(t)
    doc.build(elements)

# --- GIAO DIỆN CHÍNH ---
uploaded_file = st.file_uploader("Chọn file PDF gốc", type=["pdf"])

if uploaded_file is not None:
    # Gợi ý ngày batch
    default_date = uploaded_file.name[:6] if uploaded_file.name[:6].isdigit() else "112425"
    batch_date = st.text_input("Ngày Batch (để đặt tên file)", value=default_date)

    if st.button("🚀 Xử lý ngay"):
        with st.spinner('Đang phân tích và tách file...'):
            with tempfile.TemporaryDirectory() as temp_dir:
                # Lưu file tạm
                input_path = os.path.join(temp_dir, "input.pdf")
                with open(input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                header = []
                lo_col_index = -1
                lo_groups = {} # Dictionary để gom nhóm: {'016': [row1, row2], '045': [row3]}

                # --- BƯỚC 1: ĐỌC DỮ LIỆU TỪ TRANG 1 ---
                with pdfplumber.open(input_path) as pdf:
                    page1 = pdf.pages[0]
                    # Thử extract table
                    table = page1.extract_table()
                    
                    if table:
                        # Tìm Header và cột LO
                        header_row_idx = -1
                        for i, row in enumerate(table):
                            row_str = [str(c).strip() for c in row if c]
                            # Dấu hiệu nhận biết Header: chứa chữ "LO"
                            if "LO" in row_str:
                                header = row
                                header_row_idx = i
                                # Tìm vị trí cột LO
                                for idx, col_name in enumerate(row):
                                    if col_name and "LO" == col_name.strip():
                                        lo_col_index = idx
                                        break
                                break
                        
                        if lo_col_index != -1:
                            # Quét các dòng dữ liệu bên dưới Header
                            for row in table[header_row_idx + 1:]:
                                if row and len(row) > lo_col_index:
                                    raw_lo = row[lo_col_index]
                                    if raw_lo:
                                        # Làm sạch mã LO (bỏ xuống dòng, khoảng trắng)
                                        clean_lo = str(raw_lo).strip().replace('\n', '')
                                        
                                        # Chỉ lấy nếu LO là số (ví dụ 016, 235...)
                                        if clean_lo.isdigit():
                                            # Cập nhật lại mã sạch vào row
                                            row[lo_col_index] = clean_lo
                                            
                                            # Đưa vào nhóm (Grouping)
                                            if clean_lo not in lo_groups:
                                                lo_groups[clean_lo] = []
                                            lo_groups[clean_lo].append(row)
                        else:
                            st.error("❌ Không tìm thấy cột 'LO' trong bảng.")
                            st.stop()
                    else:
                        st.error("❌ Không đọc được bảng từ trang 1.")
                        st.stop()

                st.success(f"✅ Đã tìm thấy {len(lo_groups)} mã LO khác nhau (đã gộp các dòng trùng).")

                # --- BƯỚC 2: TẠO FILE PDF ---
                reader = PdfReader(input_path)
                total_pages = len(reader.pages)
                
                # Tạo file ZIP
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    
                    # Duyệt qua từng nhóm LO
                    for lo_id, rows in lo_groups.items():
                        pdf_name = f"{batch_date}-{lo_id}.pdf"
                        temp_page1_path = os.path.join(temp_dir, "temp_page1.pdf")
                        
                        try:
                            # A. Tạo trang 1 mới (chứa Header + danh sách rows của LO này)
                            create_page_1(rows, header, temp_page1_path, lo_id)
                            
                            # B. Ghép file
                            merger = PdfWriter()
                            
                            # 1. Thêm trang 1 vừa tạo
                            merger.add_page(PdfReader(temp_page1_path).pages[0])
                            
                            # 2. Thêm TẤT CẢ các trang còn lại từ file gốc (Từ trang 2 -> Hết)
                            if total_pages > 1:
                                for i in range(1, total_pages):
                                    merger.add_page(reader.pages[i])

                            # C. Lưu vào ZIP
                            output_pdf_buffer = BytesIO()
                            merger.write(output_pdf_buffer)
                            zip_file.writestr(pdf_name, output_pdf_buffer.getvalue())
                            
                        except Exception as e:
                            st.warning(f"Lỗi khi tạo file LO {lo_id}: {e}")

                # --- BƯỚC 3: NÚT TẢI VỀ ---
                st.download_button(
                    label=f"📥 Tải xuống {len(lo_groups)} file (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"LO_Refunds_{batch_date}.zip",
                    mime="application/zip"
                )
