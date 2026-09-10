import os
# Kiểm tra nếu đang chạy trên Windows (nt) thì mới load thư viện DLL
if os.name == 'nt':
    try:
        os.add_dll_directory(r"C:\Program Files\GTK3-Runtime Win64\bin")
    except Exception:
        pass
import io
import zipfile
from weasyprint import HTML 
import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import os
from weasyprint import HTML

# --- HÀM ĐỊNH DẠNG SỐ VNĐ ---
def format_vnd(amount):
    return f"{float(amount):,.0f}".replace(",", ".")

# --- 1. CẤU HÌNH DANH MỤC BÀ CON ---
DANH_MUC_NCC = {
    "Bè cá": ["Chú Bảy Bè Cá", "Cô Tư Bè Cá"],
    "Cá lóc bay": ["Chú Tám Cá Lóc", "Vườn Cá Lóc Chú Ba"],
    "Vườn trái cây": ["Cô Ba Vườn Trái Cây", "Vườn Chú Chín"],
    "Làm bánh dân gian": ["Dì Năm Làm Bánh", "Chị Hai Bánh Cống"],
    "Hướng dẫn viên": ["HDV Nguyễn Văn A", "HDV Trần Thị B", "HDV Lê Văn C (Tiếng Anh)"],
    "Khác": ["HTX Du lịch Sinh thái Cồn Sơn"]
}

def lay_ncc_mac_dinh(ten_dv):
    ten_lower = ten_dv.lower()
    if "bè cá" in ten_lower: return DANH_MUC_NCC["Bè cá"][0]
    if "cá lóc" in ten_lower: return DANH_MUC_NCC["Cá lóc bay"][0]
    if "trái cây" in ten_lower or "vườn" in ten_lower: return DANH_MUC_NCC["Vườn trái cây"][0]
    if "bánh" in ten_lower: return DANH_MUC_NCC["Làm bánh dân gian"][0]
    if "hdv" in ten_lower or "hướng dẫn" in ten_lower: return DANH_MUC_NCC["Hướng dẫn viên"][0]
    return DANH_MUC_NCC["Khác"][0]

# --- 2. HÀM TẠO FILE PDF (GOM NHÓM THEO BÀ CON & NGẮT TRANG) ---
def xuat_pdf_gom_nhom(ncc_name, group_df, ten_file_goc, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    htx_name = "HTX DU LỊCH NÔNG NGHIỆP CỒN SƠN"
    tong_tien = group_df["Thành tiền (VNĐ)"].sum()
    
    html_str = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            /* Định dạng trang A4 khổ đứng */
            @page {{ size: A4 portrait; margin: 20mm 15mm; }}
            body {{ font-family: 'Times New Roman', serif; font-size: 13pt; line-height: 1.5; }}
            
            .header-table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; }}
            .header-table td {{ vertical-align: top; text-align: center; }}
            .bold {{ font-weight: bold; }}
            .title {{ text-align: center; font-size: 16pt; font-weight: bold; margin: 20px 0 10px 0; }}
            
            table.data-table {{ width: 100%; border-collapse: collapse; margin: 15px 0; }}
            table.data-table th, table.data-table td {{ border: 1px solid #000; padding: 8px; text-align: left; }}
            table.data-table th {{ background-color: #f0f0f0; text-align: center; }}
            
            .signature-table {{ width: 100%; margin-top: 30px; text-align: center; }}
            .signature-table td {{ width: 50%; vertical-align: top; }}
            
            /* CSS Ép ngắt trang */
            .page-break {{ page-break-before: always; break-before: page; }}
        </style>
    </head>
    <body>
        <!-- =================== TRANG 1: BIÊN BẢN GIAO NHẬN =================== -->
        <table class="header-table">
            <tr>
                <td class="bold" style="width: 45%;">{htx_name}<br>Số: ..../BBGN-26</td>
                <td class="bold" style="width: 55%;">CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM<br>Độc lập - Tự do - Hạnh phúc</td>
            </tr>
        </table>
        <div class="title">BIÊN BẢN GIAO NHẬN DỊCH VỤ</div>
        <p>Hôm nay, đại diện <b>{htx_name}</b> và <b>{ncc_name}</b> cùng xác nhận khối lượng cung cấp dịch vụ như sau:</p>
        
        <table class="data-table">
            <tr>
                <th>STT</th>
                <th>Mã Tour</th>
                <th>Nội dung chi tiết</th>
                <th>SL</th>
                <th>Đơn giá</th>
                <th>Thành tiền</th>
            </tr>
    """
    
    # Đổ danh sách dịch vụ chi tiết vào bảng
    for i, row in enumerate(group_df.iterrows()):
        r = row[1]
        html_str += f"""
            <tr>
                <td style="text-align:center;">{i+1}</td>
                <td style="text-align:center;">{r['Mã Tour']}</td>
                <td>{r['Loại DV']}</td>
                <td style="text-align:center;">{r['SL Khách']:.0f}</td>
                <td style="text-align:right;">{format_vnd(r['Đơn giá'])}</td>
                <td style="text-align:right;">{format_vnd(r['Thành tiền (VNĐ)'])}</td>
            </tr>
        """
        
    html_str += f"""
            <tr>
                <td colspan="5" style="text-align:right;" class="bold">TỔNG CỘNG:</td>
                <td style="text-align:right;" class="bold">{format_vnd(tong_tien)}</td>
            </tr>
        </table>
        
        <table class="signature-table">
            <tr>
                <td class="bold">ĐẠI DIỆN BÊN CUNG CẤP<br><br><br><br>{ncc_name}</td>
                <td class="bold">ĐẠI DIỆN HỢP TÁC XÃ<br><br><br><br></td>
            </tr>
        </table>

        <!-- Lệnh ngắt trang sang trang thứ 2 -->
        <div class="page-break"></div>
        
        <!-- =================== TRANG 2: BIÊN BẢN THANH LÝ =================== -->
        <table class="header-table">
            <tr>
                <td class="bold" style="width: 45%;">{htx_name}<br>Số: ..../BBTL-26</td>
                <td class="bold" style="width: 55%;">CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM<br>Độc lập - Tự do - Hạnh phúc</td>
            </tr>
        </table>
        <div class="title">BIÊN BẢN THANH LÝ & THANH TOÁN</div>
        <p>Căn cứ Biên bản giao nhận dịch vụ. Hai bên đối soát và thanh toán với các nội dung sau:</p>
        <p><b>1. Giá trị thanh toán:</b> {format_vnd(tong_tien)} VNĐ</p>
        <p><b>2. Hình thức thanh toán:</b> Chuyển khoản/Tiền mặt</p>
        <p>Bên cung cấp xác nhận đã nhận đủ số tiền. Hai bên thanh lý các thỏa thuận liên quan, không còn khiếu nại về sau.</p>
        
        <table class="signature-table">
            <tr>
                <td class="bold">NGƯỜI NHẬN TIỀN<br><br><br><br>{ncc_name}</td>
                <td class="bold">ĐẠI DIỆN HỢP TÁC XÃ<br><br><br><br></td>
            </tr>
        </table>
    </body>
    </html>
    """
    
    # Định dạng tên file: BB_TenFileGoc_TenNCC.pdf
    safe_ncc_name = ncc_name.replace(" ", "_").replace("/", "")
    safe_file_goc = ten_file_goc.replace(".xml", "")
    pdf_filename = f"BB_{safe_file_goc}_{safe_ncc_name}.pdf"
    
    pdf_bytes = HTML(string=html_str).write_pdf()
    return pdf_filename, pdf_bytes
    
    return full_path

# --- 3. ĐỌC XML VÀ XỬ LÝ (Tương tự bản trước) ---
def doc_xml_va_phan_loai(file_object):
    tree = ET.parse(file_object)
    root = tree.getroot()

    def get_text(element, tag_name):
        for child in element.iter():
            clean_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if clean_tag == tag_name: return child.text
        return None

    tien_truoc_thue = float(get_text(root, 'TgTCThue') or 0.0)
    tours_can_tach = []
    tours_da_phan_bo = []
    tong_tien_items = 0

    for child in root.iter():
        clean_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if clean_tag == 'HHDVu':
            ten_dv = get_text(child, 'THHDVu')
            if not ten_dv: continue

            if '-' in ten_dv:
                ma_tour = ten_dv.split('-')[-1].strip()
                ten_goc = "-".join(ten_dv.split('-')[:-1]).strip()
            elif '/' in ten_dv:
                ma_tour = ten_dv.split('/')[-1].strip()
                ten_goc = ten_dv.split('/')[0].strip()
            else:
                ma_tour = "Không rõ"
                ten_goc = ten_dv

            try: sl = float(get_text(child, 'SLuong') or 1.0)
            except: sl = 1.0
            
            try: don_gia = float(get_text(child, 'DGia') or 0.0)
            except: don_gia = 0.0
            
            thanh_tien_item = sl * don_gia
            tong_tien_items += thanh_tien_item
            
            ten_lower = ten_dv.lower()
            tu_khoa_chi_tiet = ["bè cá", "cá lóc", "trái cây", "bánh", "hdv", "hướng dẫn"]
            
            if any(k in ten_lower for k in tu_khoa_chi_tiet):
                tours_da_phan_bo.append({
                    "Mã Tour": ma_tour,
                    "Loại DV": ten_goc,
                    "Nhà cung cấp": lay_ncc_mac_dinh(ten_goc),
                    "SL Khách": sl,
                    "Đơn giá": don_gia,
                    "Thành tiền (VNĐ)": thanh_tien_item
                })
            elif "tham quan" in ten_lower and not any(k in ten_lower for k in ["ăn", "uống"]):
                tours_can_tach.append({
                    "Mã Tour": ma_tour,
                    "Tên trên Hóa đơn": ten_dv,
                    "Số lượng khách": int(sl),
                    "Loại HDV": "Tiếng Việt"
                })
                
    if tien_truoc_thue == 0: tien_truoc_thue = tong_tien_items
    return tien_truoc_thue, pd.DataFrame(tours_can_tach), pd.DataFrame(tours_da_phan_bo)

def tinh_toan_chi_phi(df_tours_can_tach, df_tours_da_phan_bo):
    details = []
    for _, row in df_tours_can_tach.iterrows():
        ma_tour = row["Mã Tour"]
        sl = row["Số lượng khách"]
        loai_hdv = row["Loại HDV"]
        
        details.append({"Mã Tour": ma_tour, "Loại DV": "Bè cá", "Nhà cung cấp": lay_ncc_mac_dinh("Bè cá"), "SL Khách": sl, "Đơn giá": 40000})
        details.append({"Mã Tour": ma_tour, "Loại DV": "Cá lóc bay", "Nhà cung cấp": lay_ncc_mac_dinh("Cá lóc bay"), "SL Khách": sl, "Đơn giá": 30000})
        details.append({"Mã Tour": ma_tour, "Loại DV": "Vườn trái cây", "Nhà cung cấp": lay_ncc_mac_dinh("Vườn trái cây"), "SL Khách": sl, "Đơn giá": 30000})
        
        tien_banh_don_gia = 100000 if sl < 5 else 25000
        sl_banh = 1 if sl < 5 else sl
        details.append({"Mã Tour": ma_tour, "Loại DV": "Làm bánh dân gian", "Nhà cung cấp": lay_ncc_mac_dinh("bánh"), "SL Khách": sl_banh, "Đơn giá": tien_banh_don_gia})
        
        if loai_hdv == "Tiếng Việt":
            tien_hdv_don_gia = 200000 if sl < 10 else 20000
        else:
            tien_hdv_don_gia = 300000 if sl < 10 else (20000 + (100000/sl)) 
            
        sl_hdv = 1 if sl < 10 else sl
        details.append({"Mã Tour": ma_tour, "Loại DV": f"HDV {loai_hdv}", "Nhà cung cấp": lay_ncc_mac_dinh("HDV"), "SL Khách": sl_hdv, "Đơn giá": tien_hdv_don_gia})
        
    df_tach = pd.DataFrame(details)
    
    if not df_tach.empty and not df_tours_da_phan_bo.empty:
        df_final = pd.concat([df_tach, df_tours_da_phan_bo], ignore_index=True)
    elif not df_tach.empty:
        df_final = df_tach
    else:
        df_final = df_tours_da_phan_bo
        
    if not df_final.empty and "Thành tiền (VNĐ)" not in df_final.columns:
        df_final["Thành tiền (VNĐ)"] = df_final["SL Khách"] * df_final["Đơn giá"]
        
    return df_final

# --- 4. GIAO DIỆN STREAMLIT ---
st.set_page_config(page_title="Tool Xử Lý Hồ Sơ Tài Chính", layout="wide")
st.title("📄 Tool Bóc Tách Hồ Sơ Thanh Lý Du Lịch Cộng Đồng")

# Ô NHẬP ĐƯỜNG DẪN THƯ MỤC
st.sidebar.markdown("### 📂 Cấu hình lưu trữ")
thu_muc_luu = st.sidebar.text_input("Đường dẫn lưu file PDF:", value="./HoSo_Xuat_PDF", help="Copy đường dẫn thư mục trên máy tính của bạn dán vào đây (VD: D:\\DuLieu\\HoSo)")

if 'current_index' not in st.session_state:
    st.session_state.current_index = 0

uploaded_files = st.file_uploader("Tải file XML Hóa đơn lên đây", type=['xml'], accept_multiple_files=True)

if uploaded_files:
    total_files = len(uploaded_files)
    if st.session_state.current_index < total_files:
        current_file = uploaded_files[st.session_state.current_index]
        st.subheader(f"Đang xử lý ({st.session_state.current_index + 1}/{total_files}): `{current_file.name}`")
        
        tien_hd_truoc_thue, df_can_tach, df_da_phan_bo = doc_xml_va_phan_loai(current_file)
        st.info(f"💰 **Tổng tiền trên Hóa đơn (Trước Thuế):** {format_vnd(tien_hd_truoc_thue)} VNĐ")
        
        if not df_can_tach.empty:
            st.markdown("### ⚙️ Bước 1: Cấu hình Tour chung")
            loai_hdv_selections = {}
            for index, row in df_can_tach.iterrows():
                col1, col2 = st.columns([3, 1])
                with col1: st.write(f"🏷️ **{row['Tên trên Hóa đơn']}** (Đoàn {row['Số lượng khách']} khách)")
                with col2:
                    loai_hdv_selections[index] = st.selectbox("Loại HDV:", ["Tiếng Việt", "Tiếng Anh (Quốc tế)"], key=f"hdv_{index}", label_visibility="collapsed")
            for index in df_can_tach.index: df_can_tach.at[index, 'Loại HDV'] = loai_hdv_selections[index]
            
        st.markdown("### 📝 Bước 2: Bảng chi tiết chi phí (Cho phép sửa Đơn giá & Số lượng)")
        df_chi_tiet = tinh_toan_chi_phi(df_can_tach, df_da_phan_bo)
        
        if not df_chi_tiet.empty:
            tat_ca_ncc = [ncc for sublist in DANH_MUC_NCC.values() for ncc in sublist]
            
            edited_chi_tiet = st.data_editor(
                df_chi_tiet,
                column_config={
                    "Nhà cung cấp": st.column_config.SelectboxColumn("Nhà cung cấp", options=tat_ca_ncc, required=True),
                    "SL Khách": st.column_config.NumberColumn("Số lượng", format="%f"),
                    "Đơn giá": st.column_config.NumberColumn("Đơn giá (VNĐ)", format="%d"),
                    "Thành tiền (VNĐ)": st.column_config.NumberColumn("Thành tiền (VNĐ)", disabled=True)
                },
                hide_index=True,
                use_container_width=True
            )
            
            edited_chi_tiet["Thành tiền (VNĐ)"] = edited_chi_tiet["SL Khách"] * edited_chi_tiet["Đơn giá"]
            tong_chi_phi = edited_chi_tiet["Thành tiền (VNĐ)"].sum()
            chenh_lech = tien_hd_truoc_thue - tong_chi_phi
            
            col_a, col_b = st.columns(2)
            col_a.success(f"**Tổng chi trả thực tế cho Bà con:** {format_vnd(tong_chi_phi)} VNĐ")
            
            if chenh_lech != 0:
                col_b.error(f"⚠️ **Chênh lệch (HTX giữ lại/Phát sinh):** {format_vnd(chenh_lech)} VNĐ")
            else:
                col_b.success("✅ **Chênh lệch:** 0 VNĐ (Khớp 100% với Hóa đơn)")

        st.markdown("<br>", unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        
        with col1:
            # Dùng st.empty() để tạo không gian chứa nút Download sau khi xử lý xong
            download_placeholder = st.empty()
            
            if st.button("✅ Duyệt & Tạo Hồ Sơ (Gom nhóm)", use_container_width=True, type="primary"):
                with st.spinner("Đang tạo PDF và đóng gói... Team ráng đợi xíu nha!"):
                    import io
                    import zipfile
                    
                    # Tạo một bộ nhớ đệm để chứa file ZIP
                    zip_buffer = io.BytesIO()
                    
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for ncc, group_df in edited_chi_tiet.groupby("Nhà cung cấp"):
                            # Gọi hàm tạo PDF, nhận về tên file và dữ liệu nhị phân (bytes)
                            pdf_filename, pdf_bytes = xuat_pdf_gom_nhom(ncc, group_df, current_file.name, "")
                            
                            # Ghi file PDF đó vào thẳng trong cục ZIP
                            zip_file.writestr(pdf_filename, pdf_bytes)
                    
                    st.success("🎉 Đã gom xong toàn bộ biên bản!")
                    
                    # Hiện nút cho phép tải cục ZIP đó về máy
                    safe_ten_goc = current_file.name.replace('.xml', '')
                    download_placeholder.download_button(
                        label="⬇️ TẢI FILE ZIP BIÊN BẢN VỀ MÁY",
                        data=zip_buffer.getvalue(),
                        file_name=f"HoSo_ThanhLy_{safe_ten_goc}.zip",
                        mime="application/zip",
                        type="primary",
                        use_container_width=True
                    )
                
        with col2:
            if st.button("⏭️ Bỏ qua file này", use_container_width=True):
                st.session_state.current_index += 1
                st.rerun()
