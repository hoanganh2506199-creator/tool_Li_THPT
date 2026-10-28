import streamlit as st
import subprocess
import glob
import pypandoc
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from dotenv import load_dotenv
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))
from config import model  # Nạp model sau khi đọc khóa API từ .env
# 1. Hàm đọc file yêu cầu hệ thống (system.txt)
def doc_file_yeu_cau(filepath="prompts/system.txt"):
    try:
        with open(os.path.join(BASE_DIR, filepath), "r", encoding="utf-8") as file:
            return file.read()
    except FileNotFoundError:
        return "Bạn là chuyên gia biên soạn đề thi môn Vật lí cấp THPT." # Dự phòng nếu mất file

# Chuẩn hóa các đoạn chứa đủ bốn phương án A, B, C, D.
def dinh_dang_noi_dung(noi_dung):
    cac_doan = re.split(r"\n\s*\n", noi_dung.strip())
    for i, doan in enumerate(cac_doan):
        # Giữ nguyên bảng và khối mã.
        if "|" in doan or "```" in doan:
            continue
        mau = r"(?<!\S)(?:\*\*)?([ABCD])[.)](?:\*\*)?\s+"
        nhan = [m.group(1) for m in re.finditer(mau, doan)]
        if nhan == ["A", "B", "C", "D"]:
            doan = re.sub(mau, lambda m: "\n\n" + m.group(0), doan)
            cac_doan[i] = doan.strip()
    return "\n\n".join(cac_doan)


# Thêm đường kẻ cho các bảng thực sự trong file Word.
def ke_bang_word(duong_dan):
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    def w(ten):
        return "{" + ns + "}" + ten
    with zipfile.ZipFile(duong_dan, "r") as z:
        tep = {info.filename: (info, z.read(info.filename)) for info in z.infolist()}
    info, xml = tep["word/document.xml"]
    root = ET.fromstring(xml)
    for bang in root.iter(w("tbl")):
        thuoc_tinh = bang.find(w("tblPr"))
        if thuoc_tinh is None:
            thuoc_tinh = ET.Element(w("tblPr"))
            bang.insert(0, thuoc_tinh)
        vien = thuoc_tinh.find(w("tblBorders"))
        if vien is None:
            vien = ET.SubElement(thuoc_tinh, w("tblBorders"))
        for canh in ("top", "left", "bottom", "right", "insideH", "insideV"):
            node = vien.find(w(canh))
            if node is None:
                node = ET.SubElement(vien, w(canh))
            for khoa, gia_tri in {"val": "single", "sz": "4", "color": "000000"}.items():
                node.set(w(khoa), gia_tri)
    tep["word/document.xml"] = (info, ET.tostring(root, encoding="utf-8", xml_declaration=True))
    with zipfile.ZipFile(duong_dan, "w") as z:
        for info, du_lieu in tep.values():
            z.writestr(info, du_lieu)


def xuat_word_bang_pandoc(noi_dung_md, ten_file_docx="Hoàng_Anh_đề_lí_THPT.docx"):
    try:
        try:
            pypandoc.get_pandoc_version()
        except OSError:
            pypandoc.download_pandoc()
        pypandoc.convert_text(
            dinh_dang_noi_dung(noi_dung_md),
            "docx",
            format="markdown+pipe_tables+multiline_tables+grid_tables+hard_line_breaks",
            outputfile=ten_file_docx,
        )
        ke_bang_word(ten_file_docx)
        return True
    except Exception as e:
        st.error(f"❌ Lỗi khi xuất file Word: {e}")
        return False


# --- 1. DỮ LIỆU PHẠM VI KIẾN THỨC CHI TIẾT (VẬT LÍ THPT - GDPT 2018) ---
DANH_MUC_VAT_LI = {
    "Lớp 10": {
        "Chương 1: Mở đầu": ["Bài 1: Làm quen với Vật lí", "Bài 2: Các quy tắc an toàn trong phòng thực hành", "Bài 3: Thực hành tính sai số trong phép đo"],
        "Chương 2: Động học": ["Bài 4: Độ dịch chuyển và quãng đường đi được", "Bài 5: Tốc độ và vận tốc", "Bài 6: Thực hành đo tốc độ", "Bài 7: Đồ thị độ dịch chuyển - thời gian", "Bài 8: Chuyển động biến đổi. Gia tốc", "Bài 9: Chuyển động thẳng biến đổi đều", "Bài 10: Sự rơi tự do", "Bài 11: Thực hành đo gia tốc rơi tự do", "Bài 12: Chuyển động ném"],
        "Chương 3: Động lực học": ["Bài 13: Tổng hợp và phân tích lực. Cân bằng lực", "Bài 14: Định luật 1 Newton", "Bài 15: Định luật 2 Newton", "Bài 16: Định luật 3 Newton", "Bài 17: Trọng lực và lực căng", "Bài 18: Lực ma sát", "Bài 19: Lực cản và lực nâng", "Bài 20: Một số ví dụ về lực hướng tâm", "Bài 21: Moment lực. Cân bằng của vật rắn", "Bài 22: Thực hành tổng hợp lực"],
        "Chương 4: Năng lượng, công, công suất": ["Bài 23: Năng lượng. Công cơ học", "Bài 24: Công suất", "Bài 25: Động năng. Thế năng", "Bài 26: Cơ năng và định luật bảo toàn cơ năng", "Bài 27: Hiệu suất"],
        "Chương 5: Động lượng": ["Bài 28: Động lượng", "Bài 29: Định luật bảo toàn động lượng"],
        "Chương 6: Chuyển động tròn": ["Bài 31: Động học của chuyển động tròn đều", "Bài 32: Lực hướng tâm và gia tốc hướng tâm"],
        "Chương 7: Biến dạng của vật rắn. Áp suất chất lỏng": ["Bài 33: Biến dạng của vật rắn", "Bài 34: Khối lượng riêng. Áp suất chất lỏng"]
    },
    "Lớp 11": {
        "Chương 1: Dao động": ["Bài 1: Dao động điều hòa", "Bài 2: Phương trình dao động điều hòa", "Bài 3: Vận tốc, gia tốc trong dao động điều hòa", "Bài 4: Bài tập về dao động điều hòa", "Bài 5: Động năng. Thế năng. Sự chuyển hóa năng lượng", "Bài 6: Dao động tắt dần. Dao động cưỡng bức. Cộng hưởng", "Bài 7: Thực hành: Xác định chu kì dao động"],
        "Chương 2: Sóng": ["Bài 8: Mô tả sóng", "Bài 9: Sóng ngang, sóng dọc, sự truyền năng lượng của sóng cơ", "Bài 10: Thực hành: Đo tần số của sóng âm", "Bài 11: Sóng điện từ", "Bài 12: Giao thoa sóng", "Bài 13: Sóng dừng", "Bài 14: Bài tập về sóng", "Bài 15: Thực hành: Đo tốc độ truyền âm"],
        "Chương 3: Điện trường": ["Bài 16: Lực tương tác giữa hai điện tích", "Bài 17: Khái niệm điện trường", "Bài 18: Điện trường đều", "Bài 19: Thế năng điện", "Bài 20: Điện thế", "Bài 21: Tụ điện"],
        "Chương 4: Dòng điện. Mạch điện": ["Bài 22: Cường độ dòng điện", "Bài 23: Điện trở. Định luật Ohm", "Bài 24: Nguồn điện", "Bài 25: Năng lượng và công suất điện", "Bài 26: Thực hành: Đo suất điện động và điện trở trong của pin"]
    },
    "Lớp 12": {
        "Chương 1: Vật lí nhiệt": ["Bài 1: Cấu trúc của chất. Sự chuyển thể", "Bài 2: Nội năng. Định luật I Nhiệt động lực học", "Bài 3: Nhiệt độ. Thang nhiệt độ", "Bài 4: Nhiệt dung riêng", "Bài 5: Nhiệt nóng chảy riêng", "Bài 6: Nhiệt hóa hơi riêng", "Bài 7: Thực hành: Đo nhiệt dung riêng, nhiệt nóng chảy riêng"],
        "Chương 2: Khí lí tưởng": ["Bài 8: Mô hình động học phân tử chất khí", "Bài 9: Định luật Boyle", "Bài 10: Định luật Charles", "Bài 11: Phương trình trạng thái của khí lí tưởng", "Bài 12: Áp suất khí lí tưởng"],
        "Chương 3: Từ trường": ["Bài 13: Từ trường", "Bài 14: Lực từ", "Bài 15: Cảm ứng từ", "Bài 16: Từ thông. Hiện tượng cảm ứng điện từ", "Bài 17: Máy phát điện xoay chiều", "Bài 18: Ứng dụng hiện tượng cảm ứng điện từ"],
        "Chương 4: Vật lí hạt nhân": ["Bài 19: Cấu trúc hạt nhân", "Bài 20: Lực hạt nhân. Năng lượng liên kết", "Bài 21: Phóng xạ", "Bài 22: Phản ứng hạt nhân", "Bài 23: Năng lượng hạt nhân", "Bài 24: Công nghiệp hạt nhân"]
    }
}

# Cấu hình trang
st.set_page_config(page_title="Thầy giáo Hoàng Anh", layout="wide")
st.title("📚 Thầy giáo Hoàng Anh ")

# --- 2. THANH BÊN (SIDEBAR) ---
with st.sidebar:
    st.markdown("### 🎯 Cấu trúc Đề Vật lí, chọn đầy đủ các nội dung")
    
    st.markdown("**📌 Phạm vi kiến thức**")
    khoi_lop = st.selectbox("Chọn Khối lớp:", list(DANH_MUC_VAT_LI.keys()))
    
    danh_sach_chuong = list(DANH_MUC_VAT_LI[khoi_lop].keys())
    chuong_duoc_chon = st.multiselect("Chọn Chương:", danh_sach_chuong, default=[danh_sach_chuong[0]])
    
    danh_sach_bai = []
    for chuong in chuong_duoc_chon:
        danh_sach_bai.extend(DANH_MUC_VAT_LI[khoi_lop][chuong])
        
    bai_duoc_chon = st.multiselect("Chọn Bài cụ thể (Tùy chọn):", danh_sach_bai, default=danh_sach_bai)
    
    st.markdown("---")
    
    st.markdown("**1. Phân bổ Mức độ nhận thức (%):**")
    col_md1, col_md2 = st.columns(2)
    with col_md1:
        nhan_biet = st.number_input("Nhận biết (%)", min_value=0, max_value=100, value=40)
        van_dung = st.number_input("Vận dụng (%)", min_value=0, max_value=100, value=20)
    with col_md2:
        thong_hieu = st.number_input("Thông hiểu (%)", min_value=0, max_value=100, value=30)
        vd_cao = st.number_input("VD cao (%)", min_value=0, max_value=100, value=10)

    st.markdown("**2. Phân bổ Số lượng câu hỏi:**")
    col_sl1, col_sl2 = st.columns(2)
    with col_sl1:
        sl_tn_nhieu = st.number_input("Phần 1: TN Nhiều lựa chọn", min_value=0, value=18)
        sl_tn_ds = st.number_input("Phần 2: TN Đúng/Sai (x4 ý)", min_value=0, value=4)
    with col_sl2:
        sl_tl_ngan = st.number_input("Phần 3: Trả lời ngắn", min_value=0, value=6)
        sl_tu_luan = st.number_input("Phần 4: Tự luận (nếu có)", min_value=0, value=0)

# --- 3. MÀN HÌNH CHÍNH (MAIN AREA) ---
if "messages" not in st.session_state:
    st.session_state.messages = []

def lam_moi_bai():
    st.session_state.messages = []

tab_tuong_tac, tab_lich_su = st.tabs(["🚀 Môi trường làm việc", "🕰 Lịch sử đã tạo"])

with tab_tuong_tac:
    st.markdown("### 💡 1. Khởi tạo Nội dung")
    st.caption("Nhập tên bài hoặc chủ đề...")
    
    chu_de_nhap = st.text_area(
        label="Tên bài hoặc chủ đề",
        label_visibility="collapsed",
        placeholder="Ví dụ: Đề thi học kì 1 môn Vật lí 12",
        height=100
    )
    
    col_btn1, col_btn2 = st.columns([1.5, 8.5])
    with col_btn1:
        btn_bat_dau = st.button("🚀 Bắt đầu tạo", type="primary", use_container_width=True)
    with col_btn2:
        btn_lam_moi = st.button("🧹 Làm mới / Bắt đầu bài mới", on_click=lam_moi_bai)
        
    st.markdown("---")
    
    # Hiển thị lịch sử tin nhắn
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(dinh_dang_noi_dung(msg["content"]))
            
    # Xử lý nút bắt đầu tạo đề
    if btn_bat_dau:
        if not chu_de_nhap.strip():
            st.warning("Vui lòng nhập tên bài hoặc chủ đề cần tạo!")
        elif nhan_biet + thong_hieu + van_dung + vd_cao != 100:
            st.warning("Tổng tỉ lệ mức độ nhận thức phải bằng 100%.")
        elif sl_tn_nhieu + sl_tn_ds + sl_tl_ngan + sl_tu_luan == 0:
            st.warning("Vui lòng chọn ít nhất một câu hỏi.")
        else:
            yeu_cau = f"**Chủ đề:** {chu_de_nhap}"
            st.session_state.messages.append({"role": "user", "content": yeu_cau})
            
            with st.chat_message("user"):
                st.markdown(yeu_cau)
                
            with st.chat_message("assistant"):
                with st.spinner("Thầy Hoàng Anh đang tạo đề thi, vui lòng chờ, không được vội !!!"):
                    
                    chuoi_chuong = ", ".join(chuong_duoc_chon) if chuong_duoc_chon else "Toàn bộ"
                    chuoi_bai = ", ".join(bai_duoc_chon) if bai_duoc_chon else "Không giới hạn"
                    
                    # Gọi hàm đọc file system.txt
                    luat_he_thong = doc_file_yeu_cau("prompts/system.txt")
                    
                    prompt = f"""
                    {luat_he_thong}
                    
                    --- THÔNG SỐ ĐỀ BÀI LẦN NÀY ---
                    - Môn: Vật Lí THPT
                    - Khối lớp: {khoi_lop}
                    - Chủ đề cụ thể: {chu_de_nhap}
                    - Phạm vi chương: {chuoi_chuong}
                    - Giới hạn bài: {chuoi_bai}
                    
                    QUY TẮC TRÌNH BÀY BẮT BUỘC:
                    - Viết Markdown thuần, không bọc toàn bộ kết quả trong khối mã.
                    - Nội dung câu hỏi ở một dòng riêng.
                    - Mỗi phương án A., B., C., D. ở một dòng riêng;
                      cách các phương án bằng một dòng trống.
                    - Mỗi ý a), b), c), d) của câu đúng/sai ở một đoạn riêng.
                    - Ma trận và bảng đặc tả phải là bảng Markdown dùng dấu |,
                      có hàng tiêu đề và hàng phân cách | --- | --- |.
                    - Mỗi hàng bảng nằm trên một dòng; các hàng có cùng số cột.
                    - Không dùng ô gộp, khối mã, tab hoặc khoảng trắng để vẽ bảng.
                    - Có dòng trống trước và sau mỗi bảng.

                    CẤU TRÚC ĐỀ THI BẮT BUỘC:
                    - Tỉ lệ mức độ: Nhận biết {nhan_biet}%, Thông hiểu {thong_hieu}%, Vận dụng {van_dung}%, Vận dụng cao {vd_cao}%.
                    - Số lượng: {sl_tn_nhieu} câu Phần I (TN nhiều lựa chọn), {sl_tn_ds} câu Phần II (TN Đúng/Sai), {sl_tl_ngan} câu Phần III (Trả lời ngắn), {sl_tu_luan} câu tự luận.
                    """
                    
                    try:
                        response = model.generate_content(prompt)
                        ket_qua_ai = dinh_dang_noi_dung(response.text or "")

                        if ket_qua_ai and ket_qua_ai.strip():
                            st.markdown(ket_qua_ai)

                            st.session_state.messages.append({
                                "role": "assistant",
                                "content": ket_qua_ai
                            })
                        else:
                            st.warning("AI trả về nội dung trống.")
                            st.info("Kiểm tra cấu hình model và phản hồi trong Terminal.")
                            print("AI trả về nội dung trống:", response)

                    except Exception as e:
                        st.error(f"❌ Không tạo được đề: {e}")
                        st.exception(e)
                    
    # Thanh chat để tương tác với kết quả
    yeu_cau_sua = st.chat_input("Nhập yêu cầu sửa thêm (VD: cho thêm câu đồ thị, thay câu 3...)")
    
    if yeu_cau_sua:
        st.session_state.messages.append({"role": "user", "content": yeu_cau_sua})
        with st.chat_message("user"):
            st.markdown(yeu_cau_sua)
            
        with st.chat_message("assistant"):
            with st.spinner("Hoàng Anh đang đọc lại đề cũ và tinh chỉnh theo yêu cầu..."):
                
                lich_su_hoi_thoai = "Dưới đây là bối cảnh trò chuyện hiện tại:\n"
                for msg in st.session_state.messages:
                    vai_tro = "Giáo viên yêu cầu" if msg['role'] == "user" else "AI đã tạo"
                    lich_su_hoi_thoai += f"--- {vai_tro} ---\n{msg['content']}\n\n"
                
                # Gọi lại hàm đọc file system.txt để nhắc AI không quên luật
                luat_he_thong = doc_file_yeu_cau("prompts/system.txt")
                
                prompt_hoan_chinh = (f"{luat_he_thong}\n\n{lich_su_hoi_thoai}\n\n"
                    "Trả về toàn bộ đề đã sửa, mỗi phương án một đoạn riêng. "
                    "Giữ đầy đủ ma trận và bảng đặc tả dưới dạng bảng Markdown dùng dấu |.")
                
                try:
                    response = model.generate_content(prompt_hoan_chinh)
                    phan_hoi_sua = dinh_dang_noi_dung(response.text or "")
                    if not phan_hoi_sua or not phan_hoi_sua.strip():
                        raise ValueError("AI trả về nội dung trống khi sửa đề.")
                    
                    st.markdown(phan_hoi_sua)
                    st.session_state.messages.append({"role": "assistant", "content": phan_hoi_sua})
                except Exception as e:
                    st.error(f"❌ Lỗi kết nối đến AI: {e}")
                    st.exception(e)
                    
        
    # --- PHẦN MỚI THÊM: NÚT TẢI XUỐNG FILE WORD ---
    tin_nhan_ai = [msg for msg in st.session_state.messages if msg["role"] == "assistant"]
    if tin_nhan_ai:
        st.markdown("---")
        st.markdown("### 📥 Nhặt thôi nhỉ !!!")

        noi_dung_moi_nhat = tin_nhan_ai[-1]["content"]
        # Xử lý lỗi --- gây hỏng Pandoc nhưng không làm hỏng bảng
        lines = noi_dung_moi_nhat.split('\n')
        for i in range(len(lines)):
            if lines[i].strip() == '---':
                lines[i] = '***'
        noi_dung_moi_nhat = '\n'.join(lines)
        
        ten_file_word = "De_Vat_Li_Hoang_Anh.docx"
        
        thanh_cong = xuat_word_bang_pandoc(noi_dung_moi_nhat, ten_file_word)
        
        if thanh_cong:
            with open(ten_file_word, "rb") as file:
                st.download_button(
                    label="📥 TẢI FILE WORD",
                    data=file,
                    file_name=ten_file_word,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True
                )
with tab_lich_su:
    st.markdown("### 🕰 Danh sách đề thi đã lưu")
    st.info("Khu vực này hiển thị các file Word đã tạo trong các phiên làm việc trước.")