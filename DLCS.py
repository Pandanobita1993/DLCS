import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import io
import zipfile
import json
import os
from weasyprint import HTML

# --- CẤU HÌNH TRANG ---
st.set_page_config(page_title="Hệ Thống Kế Toán Cồn Sơn", layout="wide", initial_sidebar_state="expanded")

def format_vnd(amount):
    return f"{float(amount):,.0f}".replace(",", ".")

# --- 1. QUẢN LÝ DỮ LIỆU CẤU HÌNH (JSON) ---
CONFIG_FILE = "config_tour.json"

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    else:
        # Cấu hình mặc định nếu chưa có file
        default = {
            "GIA_DICH_VU": {
                "Bè cá 7 Bon": 40000,
                "Cá lóc bay": 30000,
                "Vườn trái cây": 30000,
                "Vườn Bưởi Phương My": 40000,
                "Làm bánh dân gian": 100000,
                "Xiếc Ếch & Trâu": 40000,
                "Bánh Nam Bộ 7 Món": 150000
            },
            "GOI_COMBO": {
                "gói đặc biệt": ["Bè cá 7 Bon", "Cá lóc bay", "Vườn trái cây", "Vườn Bưởi Phương My", "Làm bánh dân gian", "Xiếc Ếch & Trâu"],
                "gói đầy đủ": ["Bè cá 7 Bon", "Cá lóc bay", "Vườn trái cây", "Làm bánh dân gian", "Xiếc Ếch & Trâu"],
                "gói cơ bản 1": ["Bè cá 7 Bon", "Cá lóc bay", "Vườn trái cây", "Làm bánh dân gian"],
                "gói cơ bản 2": ["Bè cá 7 Bon", "Cá lóc bay", "Vườn trái cây", "Xiếc Ếch & Trâu"],
                "gói mini 1": ["Bè cá 7 Bon", "Xiếc Ếch & Trâu", "Vườn trái cây"],
                "gói mini 2": ["Bè cá 7 Bon", "Cá lóc bay", "Vườn trái cây"],
                "gói mini 3": ["Bè cá 7 Bon", "Bánh Nam Bộ 7 Món", "Vườn trái cây"]
            }
        }
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(default, f, ensure_ascii=False, indent=4)
        return default

def save_config(config_data):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config_data, f, ensure_ascii=False, indent=4)

config_sys = load_config()

# --- 2. DANH MỤC NHÀ VƯỜN ---
DANH_MUC_NCC = {
    "Bè cá": ["Chú Bảy Bè Cá", "Cô Tư Bè Cá"],
    "Cá lóc bay": ["Chú Tám Cá Lóc", "Vườn Cá Lóc Chú Ba"],
    "Vườn trái cây": [
        "Cô Ba Vườn Trái Cây", "Vườn Mận Cẩm Thanh", "Vườn Bưởi Phương My", 
        "Vườn Nhãn Xuồng Gia Khang", "Vườn Nhãn Xuồng Công Minh", 
        "Vườn Chôm Chôm Song Khánh", "Vườn Nhãn Long Hiếu Trang"
    ],
    "Làm bánh dân gian": ["Dì Năm Làm Bánh", "Chị Hai Bánh Cống", "Điểm Làm Cốm Nổ", "Điểm Đổ Bánh Xèo", "Điểm Làm Bánh Canh"],
    "Hướng dẫn viên": ["HDV Nguyễn Văn A", "HDV Trần Thị B", "HDV Lê Văn C (Tiếng Anh)"],
    "Khác": ["HTX Du lịch Sinh thái Cồn Sơn", "Điểm Xiếc Ếch & Trâu Độc Lạ"]
}

def lay_ncc_mac_dinh(ten_dv):
    ten_lower = ten_dv.lower()
    if "bè cá" in ten_lower: return DANH_MUC_NCC["Bè cá"][0]
    if "cá lóc" in ten_lower: return DANH_MUC_NCC["Cá lóc bay"][0]
    if "trái cây" in ten_lower or "bưởi" in ten_lower or "nhãn" in ten_lower or "mận" in ten_lower: return DANH_MUC_NCC["Vườn trái cây"][0]
    if "bánh" in ten_lower or "cốm" in ten_lower or "xèo" in ten_lower: return DANH_MUC_NCC["Làm bánh dân gian"][0]
    if "hdv" in ten_lower or "hướng dẫn" in ten_lower: return DANH_MUC_NCC["Hướng dẫn viên"][0]
    if "xiếc" in ten_lower or "trâu" in ten_lower: return DANH_MUC_NCC["Khác"][1]
    return DANH_MUC_NCC["Khác"][0]

# --- 3. XUẤT PDF ---
def xuat_pdf_gom_nhom(ncc_name, group_df, ten_file_goc):
    htx_name = "HTX DU LỊCH NÔNG NGHIỆP CỒN SƠN"
    tong_tien = group_df["Thành tiền (VNĐ)"].sum()
    
    html_str = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
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
            .page-break {{ page-break-before: always; break-before: page; }}
        </style>
    </head>
    <body>
        <table class="header-table">
            <tr>
                <td class="bold" style="width: 45%;">{htx_name}<br>Số: ..../BBGN-26</td>
                <td class="bold" style="width: 55%;">CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM<br>Độc lập - Tự do - Hạnh phúc</td>
            </tr>
        </table>
        <div class="title">BIÊN BẢN GIAO NHẬN DỊCH VỤ</div>
        <p>Hôm nay, đại diện <b>{htx_name}</b> và <b>{ncc_name}</b> cùng xác nhận khối lượng cung cấp dịch vụ như sau:</p>
        <table class="data-table">
            <tr><th>STT</th><th>Mã Tour</th><th>Nội dung chi tiết</th><th>SL</th><th>Đơn giá</th><th>Thành tiền</th></tr>
    """
    for i, row in enumerate(group_df.iterrows()):
        r = row[1]
        html_str += f"<tr><td style='text-align:center;'>{i+1}</td><td style='text-align:center;'>{r['Mã Tour']}</td><td>{r['Loại DV']}</td><td style='text-align:center;'>{r['SL Khách']:.0f}</td><td style='text-align:right;'>{format_vnd(r['Đơn giá'])}</td><td style='text-align:right;'>{format_vnd(r['Thành tiền (VNĐ)'])}</td></tr>"
    html_str += f"<tr><td colspan='5' style='text-align:right;' class='bold'>TỔNG CỘNG:</td><td style='text-align:right;' class='bold'>{format_vnd(tong_tien)}</td></tr></table>"
    html_str += f"<table class='signature-table'><tr><td class='bold'>ĐẠI DIỆN BÊN CUNG CẤP<br><br><br><br>{ncc_name}</td><td class='bold'>ĐẠI DIỆN HỢP TÁC XÃ<br><br><br><br></td></tr></table>"
    html_str += "<div class='page-break'></div>"
    
    html_str += f"""
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
        <table class="signature-table"><tr><td class="bold">NGƯỜI NHẬN TIỀN<br><br><br><br>{ncc_name}</td><td class="bold">ĐẠI DIỆN HỢP TÁC XÃ<br><br><br><br></td></tr></table>
    </body></html>
    """
    safe_ncc_name = ncc_name.replace(" ", "_").replace("/", "")
    safe_file_goc = ten_file_goc.replace(".xml", "")
    pdf_filename = f"BB_{safe_file_goc}_{safe_ncc_name}.pdf"
    
    return pdf_filename, HTML(string=html_str).write_pdf()

# --- 4. BÓC TÁCH XML ---
def doc_xml_va_phan_loai(file_object):
    tree = ET.parse(file_object)
    root = tree.getroot()

    def get_text(element, tag_name):
        for child in element.iter():
            clean_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if clean_tag == tag_name: return child.text
        return None

    tien_truoc_thue = float(get_text(root, 'TgTCThue') or 0.0)
    tours_can_tach, tours_da_phan_bo, tong_tien_items = [], [], 0

    for child in root.iter():
        clean_tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if clean_tag == 'HHDVu':
            ten_dv = get_text(child, 'THHDVu')
            if not ten_dv: continue

            ma_tour = ten_dv.split('-')[-1].strip() if '-' in ten_dv else (ten_dv.split('/')[-1].strip() if '/' in ten_dv else "Không rõ")
            ten_goc = "-".join(ten_dv.split('-')[:-1]).strip() if '-' in ten_dv else (ten_dv.split('/')[0].strip() if '/' in ten_dv else ten_dv)

            try: sl = float(get_text(child, 'SLuong') or 1.0)
            except: sl = 1.0
            
            try: don_gia = float(get_text(child, 'DGia') or 0.0)
            except: don_gia = 0.0
            
            tong_tien_items += (sl * don_gia)
            ten_lower = ten_dv.lower()
            
            if any(k in ten_lower for k in ["bè cá", "cá lóc", "trái cây", "bánh", "hdv", "cốm", "xèo", "canh", "xiếc", "trâu", "bưởi"]):
                tours_da_phan_bo.append({
                    "Mã Tour": ma_tour, "Loại DV": ten_goc, "Nhà cung cấp": lay_ncc_mac_dinh(ten_goc),
                    "SL Khách": sl, "Đơn giá": don_gia, "Thành tiền (VNĐ)": sl * don_gia
                })
            elif "tham quan" in ten_lower or "gói" in ten_lower:
                tours_can_tach.append({"Mã Tour": ma_tour, "Tên trên Hóa đơn": ten_dv, "Số lượng khách": int(sl)})
                
    if tien_truoc_thue == 0: tien_truoc_thue = tong_tien_items
    return tien_truoc_thue, pd.DataFrame(tours_can_tach), pd.DataFrame(tours_da_phan_bo)

# TÍNH TOÁN DỰA TRÊN FILE CẤU HÌNH (JSON)
def tinh_toan_chi_phi(df_tours_can_tach, df_tours_da_phan_bo, config):
    details = []
    for _, row in df_tours_can_tach.iterrows():
        ma_tour = row["Mã Tour"]
        sl = row["Số lượng khách"]
        ten_tour = row["Tên trên Hóa đơn"].lower()
        
        loai_hdv = row.get("Loại HDV", "Tiếng Việt")
        nha_vuon = row.get("Nhà Vườn", lay_ncc_mac_dinh("Vườn trái cây"))
        diem_banh = row.get("Làm Bánh", lay_ncc_mac_dinh("Làm bánh dân gian"))
        
        matched_combo = None
        for goi_name, t_phan in config["GOI_COMBO"].items():
            if goi_name in ten_tour:
                matched_combo = t_phan
                break
                
        # Nếu không khớp Gói nào, xé lẻ theo Nhóm 1 mặc định
        if not matched_combo:
            matched_combo = ["Bè cá 7 Bon", "Cá lóc bay", "Vườn trái cây", "Làm bánh dân gian"]
            
        for dv_name in matched_combo:
            price = config["GIA_DICH_VU"].get(dv_name, 0)
            
            # Logic riêng cho Bánh (Giá sỉ <5 khách)
            sl_dv = sl
            if "bánh" in dv_name.lower() or "cốm" in dv_name.lower():
                if sl < 5:
                    sl_dv = 1
                    price = 150000 if "7 món" in dv_name.lower() else 100000
                else:
                    price = 30000 if "7 món" in dv_name.lower() else 25000

            ncc = lay_ncc_mac_dinh(dv_name)
            if any(k in dv_name.lower() for k in ["trái cây", "bưởi", "nhãn", "mận", "chôm"]): ncc = nha_vuon
            if any(k in dv_name.lower() for k in ["bánh", "cốm", "xèo"]): ncc = diem_banh
            
            details.append({"Mã Tour": ma_tour, "Loại DV": dv_name, "Nhà cung cấp": ncc, "SL Khách": sl_dv, "Đơn giá": price})

        tien_hdv_don_gia = (200000 if sl < 10 else 20000) if loai_hdv == "Tiếng Việt" else (300000 if sl < 10 else (20000 + (100000/sl)))
        details.append({"Mã Tour": ma_tour, "Loại DV": f"HDV {loai_hdv}", "Nhà cung cấp": lay_ncc_mac_dinh("HDV"), "SL Khách": (1 if sl < 10 else sl), "Đơn giá": tien_hdv_don_gia})
        
    df_tach = pd.DataFrame(details)
    df_final = pd.concat([df_tach, df_tours_da_phan_bo], ignore_index=True) if not df_tach.empty and not df_tours_da_phan_bo.empty else (df_tach if not df_tach.empty else df_tours_da_phan_bo)
    if not df_final.empty and "Thành tiền (VNĐ)" not in df_final.columns: df_final["Thành tiền (VNĐ)"] = df_final["SL Khách"] * df_final["Đơn giá"]
    return df_final


# ==========================================
# PHẦN 5: ĐĂNG NHẬP VÀ ĐIỀU HƯỚNG GIAO DIỆN
# ==========================================

if 'logged_in' not in st.session_state: st.session_state.logged_in = False
def check_login(username, password): return username == "admin" and password == "conson2026"

if not st.session_state.logged_in:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<h2 style='text-align: center;'>🔐 Đăng nhập Hệ thống</h2>", unsafe_allow_html=True)
        if st.button("🚀 Đăng nhập", use_container_width=True, type="primary"):
            if check_login(st.text_input("Tên đăng nhập"), st.text_input("Mật khẩu", type="password")):
                st.session_state.logged_in = True
                st.rerun()
            else: st.error("Sai tài khoản/mật khẩu!")
else:
    st.sidebar.markdown("### 👤 Quản trị viên")
    if st.sidebar.button("🚪 Đăng xuất"):
        st.session_state.logged_in = False
        st.rerun()
        
    menu = st.sidebar.radio("📌 ĐIỀU HƯỚNG", ["📑 Xử lý Hóa đơn XML", "⚙️ Cài đặt Gói Tour (Admin)"])
    
    # ---------------------------------------------------------
    # TRANG 1: XỬ LÝ HÓA ĐƠN
    # ---------------------------------------------------------
    if menu == "📑 Xử lý Hóa đơn XML":
        st.title("📄 Bóc Tách Hồ Sơ Thanh Lý - Cồn Sơn")
        if 'current_index' not in st.session_state: st.session_state.current_index = 0
        uploaded_files = st.file_uploader("Tải file XML Hóa đơn", type=['xml'], accept_multiple_files=True)

        if uploaded_files and st.session_state.current_index < len(uploaded_files):
            current_file = uploaded_files[st.session_state.current_index]
            tien_hd_truoc_thue, df_can_tach, df_da_phan_bo = doc_xml_va_phan_loai(current_file)
            st.info(f"💰 **Hóa đơn:** {current_file.name} | **Tổng tiền (Trước Thuế):** {format_vnd(tien_hd_truoc_thue)} VNĐ")
            
            if not df_can_tach.empty:
                st.markdown("### ⚙️ Bước 1: Gán Nhà Vườn & HDV")
                hdv_sel, vuon_sel, banh_sel = {}, {}, {}
                for index, row in df_can_tach.iterrows():
                    st.markdown(f"**🏷️ {row['Tên trên Hóa đơn']}** *(Đoàn {row['Số lượng khách']} khách)*")
                    c1, c2, c3 = st.columns(3)
                    with c1: hdv_sel[index] = st.selectbox("HDV:", ["Tiếng Việt", "Tiếng Anh (Quốc tế)"], key=f"hdv_{index}")
                    with c2: vuon_sel[index] = st.selectbox("Vườn Trái Cây:", DANH_MUC_NCC["Vườn trái cây"], key=f"v_{index}")
                    with c3: banh_sel[index] = st.selectbox("Điểm Làm Bánh:", DANH_MUC_NCC["Làm bánh dân gian"], key=f"b_{index}")
                for idx in df_can_tach.index:
                    df_can_tach.at[idx, 'Loại HDV'] = hdv_sel[idx]
                    df_can_tach.at[idx, 'Nhà Vườn'] = vuon_sel[idx]
                    df_can_tach.at[idx, 'Làm Bánh'] = banh_sel[idx]
                
            st.markdown("### 📝 Bước 2: Bảng chi tiết chi phí")
            df_chi_tiet = tinh_toan_chi_phi(df_can_tach, df_da_phan_bo, config_sys)
            
            if not df_chi_tiet.empty:
                tat_ca_ncc = [ncc for sublist in DANH_MUC_NCC.values() for ncc in sublist]
                edited_chi_tiet = st.data_editor(
                    df_chi_tiet,
                    column_config={
                        "Nhà cung cấp": st.column_config.SelectboxColumn("Nhà cung cấp", options=tat_ca_ncc),
                        "Đơn giá": st.column_config.NumberColumn("Đơn giá", format="%d")
                    },
                    hide_index=True, use_container_width=True
                )
                edited_chi_tiet["Thành tiền (VNĐ)"] = edited_chi_tiet["SL Khách"] * edited_chi_tiet["Đơn giá"]
                tong_chi_phi = edited_chi_tiet["Thành tiền (VNĐ)"].sum()
                
                c_a, c_b = st.columns(2)
                c_a.success(f"**Chi trả cho Bà con:** {format_vnd(tong_chi_phi)} VNĐ")
                chenh_lech = tien_hd_truoc_thue - tong_chi_phi
                if chenh_lech != 0: c_b.error(f"⚠️ **Chênh lệch (HTX giữ):** {format_vnd(chenh_lech)} VNĐ")
                
                c1, c2 = st.columns(2)
                with c1:
                    download_ph = st.empty()
                    if st.button("✅ Duyệt & Tạo Hồ Sơ", type="primary", use_container_width=True):
                        with st.spinner("Đang gom biên bản..."):
                            z_buf = io.BytesIO()
                            with zipfile.ZipFile(z_buf, "w", zipfile.ZIP_DEFLATED) as z_f:
                                for ncc, grp in edited_chi_tiet.groupby("Nhà cung cấp"):
                                    pdf_fn, pdf_b = xuat_pdf_gom_nhom(ncc, grp, current_file.name)
                                    z_f.writestr(pdf_fn, pdf_b)
                            download_ph.download_button("⬇️ TẢI FILE ZIP VỀ MÁY", z_buf.getvalue(), f"HoSo_{current_file.name}.zip", "application/zip", type="primary", use_container_width=True)
                with c2:
                    if st.button("⏭️ Bỏ qua", use_container_width=True):
                        st.session_state.current_index += 1; st.rerun()

    # ---------------------------------------------------------
    # TRANG 2: QUẢN LÝ GÓI TOUR (ADMIN SETTINGS)
    # ---------------------------------------------------------
    elif menu == "⚙️ Cài đặt Gói Tour (Admin)":
        st.title("⚙️ Bảng Điều Khiển: Cấu hình Giá & Gói Tour")
        st.info("Bất kỳ thay đổi nào ở đây sẽ tự động cập nhật vào thuật toán bóc tách hóa đơn.")
        
        # BẢNG 1: GIÁ DỊCH VỤ
        st.markdown("### 1. Bảng Giá Dịch Vụ Cố Định (VNĐ)")
        df_gia = pd.DataFrame(list(config_sys["GIA_DICH_VU"].items()), columns=["Tên Dịch Vụ", "Đơn Giá Gốc"])
        edited_gia = st.data_editor(df_gia, num_rows="dynamic", use_container_width=True)
        
        # BẢNG 2: CẤU HÌNH COMBO
        st.markdown("### 2. Định Nghĩa Gói Tour (Combo)")
        st.write("*(Nhập các dịch vụ cấu thành gói cách nhau bằng dấu phẩy `,`. Tên dịch vụ phải gõ đúng y chang Bảng 1)*")
        combo_list = [{"Tên Gói Trên Hóa Đơn": k.title(), "Các Dịch Vụ Bao Gồm": ", ".join(v)} for k, v in config_sys["GOI_COMBO"].items()]
        df_combo = pd.DataFrame(combo_list)
        edited_combo = st.data_editor(df_combo, num_rows="dynamic", use_container_width=True)
        
        if st.button("💾 LƯU CẤU HÌNH VÀO HỆ THỐNG", type="primary"):
            new_gia = {row["Tên Dịch Vụ"]: row["Đơn Giá Gốc"] for _, row in edited_gia.iterrows() if str(row["Tên Dịch Vụ"]).strip()}
            new_combo = {str(row["Tên Gói Trên Hóa Đơn"]).lower().strip(): [x.strip() for x in str(row["Các Dịch Vụ Bao Gồm"]).split(",")] for _, row in edited_combo.iterrows() if str(row["Tên Gói Trên Hóa Đơn"]).strip()}
            
            config_sys["GIA_DICH_VU"] = new_gia
            config_sys["GOI_COMBO"] = new_combo
            save_config(config_sys)
            st.success("✅ Đã lưu quy tắc chia tiền mới! Hãy quay lại trang Xử Lý Hóa Đơn để test thử.")
