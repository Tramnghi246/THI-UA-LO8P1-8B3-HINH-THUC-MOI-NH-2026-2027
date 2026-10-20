import io
import os
import time
import pandas as pd
import requests
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# 1. CẤU HÌNH TRANG & GIAO DIỆN CHUẨN CỔNG THÔNG TIN TRƯỜNG THCS CHÍ CÔNG
# ==============================================================================
st.set_page_config(
    page_title="Hệ Thống Thi Đua Lớp 8B3 - THCS Chí Công",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Thay đường dẫn link web app /exec từ Google Apps Script của bạn vào đây nếu có
API_URL = "https://script.google.com/macros/s/AKfycbwoy6GT8XEkwwrLTSXrIk3tMj6lzATeiwrw02cWYz7yEXDhzEtDOoyL8VssgXuJE08q/exec"
DATA_FILE = "du_lieu_thi_dua.csv"

# ==============================================================================
# 2. BỘ CSS GIAO DIỆN HIỆN ĐẠI (SIDEBAR THẺ BO TRÒN & BANNER TRƯỜNG HỌC)
# ==============================================================================
st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Segoe+UI:wght@400;600;700&display=swap');
        html, body, [class*="css"] {
            font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
        }
        .main {
            background-color: #f4f6f9;
        }

        /* Nền và khung Sidebar */
        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #f8fafc 0%, #edf2f7 100%) !important;
            border-right: 1px solid #e2e8f0;
            box-shadow: 3px 0 12px rgba(0,0,0,0.03);
        }
        
        /* Card nhận diện thương hiệu trường trong Sidebar */
        .sidebar-brand-card {
            background: #ffffff;
            border-radius: 12px;
            padding: 14px 16px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 2px 6px rgba(11, 79, 138, 0.08);
            margin-bottom: 18px;
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .sidebar-brand-icon {
            width: 44px;
            height: 44px;
            background: #0b4f8a;
            color: #ffffff;
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.4rem;
        }
        .sidebar-brand-title {
            font-size: 1.05rem;
            font-weight: 800;
            color: #0b4f8a;
            line-height: 1.2;
        }
        .sidebar-brand-sub {
            font-size: 0.78rem;
            color: #64748b;
            font-weight: 500;
        }

        /* Menu điều hướng dạng thẻ bo tròn (ẩn radio tròn trơn) */
        div[data-testid="stRadio"] > div {
            gap: 8px !important;
        }
        div[data-testid="stRadio"] label {
            background: #ffffff !important;
            border: 1px solid #e2e8f0 !important;
            border-radius: 10px !important;
            padding: 10px 14px !important;
            margin-bottom: 2px !important;
            cursor: pointer !important;
            transition: all 0.2s ease !important;
            box-shadow: 0 1px 3px rgba(0,0,0,0.02) !important;
            display: flex !important;
            align-items: center !important;
        }
        div[data-testid="stRadio"] label > div:first-child {
            display: none !important;
        }
        div[data-testid="stRadio"] label p {
            font-size: 0.92rem !important;
            font-weight: 600 !important;
            color: #334155 !important;
            margin: 0 !important;
        }
        div[data-testid="stRadio"] label:hover {
            background-color: #f0f7ff !important;
            border-color: #93c5fd !important;
            transform: translateX(4px);
        }
        div[data-testid="stRadio"] label[data-checked="true"],
        div[data-testid="stRadio"] label:has(input:checked) {
            background: #0b4f8a !important;
            border-color: #0b4f8a !important;
            box-shadow: 0 3px 8px rgba(11, 79, 138, 0.25) !important;
        }
        div[data-testid="stRadio"] label[data-checked="true"] p,
        div[data-testid="stRadio"] label:has(input:checked) p {
            color: #ffffff !important;
            font-weight: 700 !important;
        }

        .sidebar-section-label {
            font-size: 0.75rem;
            text-transform: uppercase;
            letter-spacing: 0.6px;
            font-weight: 700;
            color: #64748b;
            margin: 14px 0 6px 4px;
            display: flex;
            align-items: center;
            gap: 6px;
        }

        /* Banner trường học ở đầu nội dung */
        .school-top-bar {
            background: linear-gradient(90deg, #0b4f8a 0%, #1565c0 100%);
            color: #ffffff;
            padding: 12px 20px;
            border-radius: 10px;
            margin-bottom: 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 2px 8px rgba(11, 79, 138, 0.2);
        }
        .school-title {
            font-size: 1.05rem;
            font-weight: 700;
            letter-spacing: 0.5px;
        }
        .school-slogan {
            font-size: 0.82rem;
            color: #fde047;
            font-weight: 600;
        }

        /* Thẻ hồ sơ học sinh */
        .student-profile-card {
            background: #ffffff;
            border-radius: 12px;
            padding: 16px 20px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 18px;
            flex-wrap: wrap;
            gap: 12px;
        }
        .student-avatar {
            width: 48px;
            height: 48px;
            background-color: #0b4f8a;
            color: #ffffff;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: 700;
            font-size: 1.25rem;
        }

        /* Thẻ thống kê */
        .stat-card {
            background: white;
            padding: 16px 20px;
            border-radius: 12px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.05);
            border: 1px solid #e2e8f0;
            display: flex;
            flex-direction: column;
            gap: 4px;
        }
        .stat-label {
            font-size: 12px;
            color: #64748b;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .stat-value {
            font-size: 26px;
            font-weight: 700;
            color: #0f172a;
        }

        /* Nút bấm và hộp thông báo */
        .stButton>button {
            border-radius: 8px;
            font-weight: 600;
            transition: all 0.2s ease;
        }
        .stButton>button:hover {
            transform: translateY(-1px);
            box-shadow: 0 4px 10px rgba(0,0,0,0.12);
        }
        .empty-log-box {
            background: white;
            border: 2px dashed #cbd5e1;
            border-radius: 12px;
            padding: 28px;
            text-align: center;
            margin: 15px 0;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==============================================================================
# 3. DANH MỤC THI ĐUA & DỮ LIỆU KHỞI TẠO
# ==============================================================================
DANH_SACH_LOI = {
    "🌸 Thầy cô tuyên dương (+2)": 2,
    "🌟 Phát biểu xây dựng bài (+1)": 1,
    "💯 Đạt điểm 9, 10 (+2)": 2,
    "🏆 Đạt giải phong trào (+5)": 5,
    "🤬 Nói tục (-5)": -5,
    "📱 Làm việc riêng (-5)": -5,
    "🔊 Nói chuyện trong giờ (-2)": -2,
    "✉️ Nghỉ học có phép (-2)": -2,
    "⏰ Đi muộn (-1)": -1,
    "📚 Không học bài / Thiếu BTVN (-5)": -5,
    "👔 Không mặc đồng phục (-5)": -5,
    "🧹 Không trực nhật / vệ sinh (-5)": -5,
    "📱 Sử dụng điện thoại không xin phép (-15)": -15,
    "🚨 Nghỉ học không phép (-5)": -5,
}

# Khung danh sách học sinh mẫu theo 4 tổ
DEFAULT_STUDENTS = [
    {"STT": 1, "Họ và tên": "Huỳnh Bảo Ngọc Thiên Ân", "Tổ": "Tổ 1", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 2, "Họ và tên": "Lê Minh Thái Dương", "Tổ": "Tổ 1", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 3, "Họ và tên": "Ngô Gia Hân", "Tổ": "Tổ 1", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 4, "Họ và tên": "Phạm Hoàng Đăng Khôi", "Tổ": "Tổ 1", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 5, "Họ và tên": "Trần Thị Cẩm Ly", "Tổ": "Tổ 1", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 6, "Họ và tên": "Vũ Đức Trọng", "Tổ": "Tổ 1", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 7, "Họ và tên": "Nguyễn Tuyền An", "Tổ": "Tổ 2", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 8, "Họ và tên": "Nguyễn Thị Ngọc Hiên", "Tổ": "Tổ 2", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 9, "Họ và tên": "Huỳnh Ngọc Huy", "Tổ": "Tổ 2", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 10, "Họ và tên": "Nguyễn Huỳnh Gia Hưng", "Tổ": "Tổ 2", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 11, "Họ và tên": "Trần Quốc Bảo", "Tổ": "Tổ 2", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 12, "Họ và tên": "Nguyễn Minh Khang", "Tổ": "Tổ 2", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 13, "Họ và tên": "Bùi Thị Tuyết Mai", "Tổ": "Tổ 2", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 14, "Họ và tên": "Lê Tuấn Hưng", "Tổ": "Tổ 3", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 15, "Họ và tên": "Lê Quỳnh Diễm Hương", "Tổ": "Tổ 3", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 16, "Họ và tên": "Nguyễn Lê Hoàng Mai", "Tổ": "Tổ 3", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 17, "Họ và tên": "Đỗ Hoàng Nam", "Tổ": "Tổ 3", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 18, "Họ và tên": "Võ Thị Yến Nhi", "Tổ": "Tổ 3", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 19, "Họ và tên": "Phan Thanh Phong", "Tổ": "Tổ 3", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 20, "Họ và tên": "Võ Huỳnh Minh An", "Tổ": "Tổ 4", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 21, "Họ và tên": "Lê Văn Tuấn", "Tổ": "Tổ 4", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 22, "Họ và tên": "Đặng Ngọc Như Quỳnh", "Tổ": "Tổ 4", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 23, "Họ và tên": "Nguyễn Tấn Tài", "Tổ": "Tổ 4", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 24, "Họ và tên": "Lê Thị Thảo Trang", "Tổ": "Tổ 4", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""},
    {"STT": 25, "Họ và tên": "Nguyễn Hoàng Việt", "Tổ": "Tổ 4", "Điểm thi đua": 100, "Xếp loại": "Tốt", "Lỗi vi phạm": ""}
]

# ==============================================================================
# 4. HÀM XỬ LÝ DỮ LIỆU
# ==============================================================================
def tinh_xep_loai(diem):
    try:
        diem = float(diem)
    except Exception:
        diem = 100.0
    if diem >= 100:
        return "Tốt"
    elif diem >= 85:
        return "Khá"
    elif diem >= 70:
        return "Trung bình"
    else:
        return "Yếu"

def parse_item_info(item_str):
    diem_goc = 0
    matched_loi = list(DANH_SACH_LOI.keys())[0]
    so_lan = 1
    ghi_chu = ""

    if "[" in item_str and "]" in item_str:
        ghi_chu = item_str[item_str.find("[") + 1 : item_str.rfind("]")].strip()
    elif "(" in item_str and ")" in item_str and not any(k in item_str for k in ["(+", "(-"]):
        ghi_chu = item_str[item_str.find("(") + 1 : item_str.rfind(")")].strip()

    for ten_loi, diem in DANH_SACH_LOI.items():
        ten_ngan = ten_loi.split(" (")[0]
        if item_str.startswith(ten_ngan):
            matched_loi = ten_loi
            diem_goc = diem
            if " x" in item_str:
                try:
                    so_lan = int(item_str.split(" x")[1].split(" ")[0].split("[")[0].split("(")[0])
                except Exception:
                    so_lan = 1
            break

    total_delta = diem_goc * so_lan
    return matched_loi, so_lan, ghi_chu, total_delta

def clean_dataframe(df_input):
    if df_input is None or df_input.empty:
        return pd.DataFrame(DEFAULT_STUDENTS)

    col_found_loi = None
    for col in df_input.columns:
        c_low = str(col).strip().lower()
        if any(k in c_low for k in ["lỗi", "vi phạm", "ghi chú", "tuyên dương", "chi tiết", "nội dung", "log"]):
            col_found_loi = col
            break

    rename_dict = {}
    for col in df_input.columns:
        c_low = str(col).strip().lower()
        if c_low in ["họ và tên", "họ tên", "tên học sinh", "ten"]:
            rename_dict[col] = "Họ và tên"
        elif c_low in ["tổ", "to", "nhóm"]:
            rename_dict[col] = "Tổ"
        elif c_low in ["điểm thi đua", "điểm", "diem", "diem thi dua"]:
            rename_dict[col] = "Điểm thi đua"
        elif c_low in ["xếp loại", "xep loai", "loại"]:
            rename_dict[col] = "Xếp loại"

    if rename_dict:
        df_input = df_input.rename(columns=rename_dict)

    if col_found_loi and col_found_loi in df_input.columns and col_found_loi != "Lỗi vi phạm":
        df_input["Lỗi vi phạm"] = df_input[col_found_loi]

    if "Lỗi vi phạm" not in df_input.columns:
        if len(df_input.columns) >= 6:
            df_input["Lỗi vi phạm"] = df_input.iloc[:, 5]
        else:
            df_input["Lỗi vi phạm"] = ""

    df_clean = df_input.dropna(subset=["Họ và tên"]).copy()
    df_clean["Họ và tên"] = df_clean["Họ và tên"].astype(str).str.strip()
    df_clean = df_clean[~df_clean["Họ và tên"].str.lower().isin(["nan", "none", ""])]

    df_clean["Tổ"] = df_clean["Tổ"].fillna("Tổ 1").astype(str).str.strip()
    df_clean["Điểm thi đua"] = (
        pd.to_numeric(df_clean["Điểm thi đua"], errors="coerce").fillna(100).astype(int)
    )
    df_clean["Xếp loại"] = df_clean["Điểm thi đua"].apply(tinh_xep_loai)
    
    df_clean["Lỗi vi phạm"] = df_clean["Lỗi vi phạm"].astype(str).str.strip()
    df_clean["Lỗi vi phạm"] = df_clean["Lỗi vi phạm"].replace(["nan", "None", "NaN", "null"], "")

    df_clean = df_clean.reset_index(drop=True)
    df_clean["STT"] = list(range(1, len(df_clean) + 1))
    return df_clean

def load_data():
    if API_URL and "/exec" in API_URL:
        try:
            res = requests.get(API_URL, params={"_t": time.time()}, timeout=7)
            if res.status_code == 200:
                data = res.json()
                if isinstance(data, list) and len(data) > 1:
                    df_temp = pd.DataFrame(data[1:], columns=data[0])
                    df_temp = clean_dataframe(df_temp)
                    df_temp.to_csv(DATA_FILE, index=False)
                    return df_temp
        except Exception:
            pass

    if os.path.exists(DATA_FILE):
        try:
            return clean_dataframe(pd.read_csv(DATA_FILE))
        except Exception:
            pass

    return pd.DataFrame(DEFAULT_STUDENTS)

def update_gsheet(name, score, xeploai, log_text):
    if not API_URL or "/exec" not in API_URL:
        return False, "Chưa thiết lập API Google Sheets"
    try:
        params = {
            "action": "update",
            "name": name,
            "score": score,
            "xeploai": xeploai,
            "log": log_text,
            "_t": time.time(),
        }
        res = requests.get(API_URL, params=params, timeout=10)
        if res.status_code == 200:
            res_json = res.json()
            if res_json.get("status") == "success":
                return True, "Thành công"
            return False, res_json.get("message", "Lỗi dữ liệu Google Apps Script")
        return False, f"Lỗi HTTP {res.status_code}"
    except Exception as e:
        return False, str(e)

def reset_new_week_gsheet():
    if not API_URL or "/exec" not in API_URL:
        return False, "Chưa thiết lập API Google Sheets"
    try:
        params = {"action": "reset_week", "_t": time.time()}
        res = requests.get(API_URL, params=params, timeout=12)
        if res.status_code == 200:
            res_json = res.json()
            if res_json.get("status") == "success":
                return True, "Thành công"
            return False, res_json.get("message", "Lỗi reset")
        return False, f"Lỗi HTTP {res.status_code}"
    except Exception as e:
        return False, str(e)

if "df" not in st.session_state:
    st.session_state.df = load_data()

df = st.session_state.df

# ==============================================================================
# 5. THANH ĐIỀU HƯỚNG BÊN TRÁI (SIDEBAR)
# ==============================================================================
with st.sidebar:
    st.markdown(
        """
        <div class="sidebar-brand-card">
            <div class="sidebar-brand-icon">🏫</div>
            <div>
                <div class="sidebar-brand-title">THCS CHÍ CÔNG</div>
                <div class="sidebar-brand-sub">Quản lý thi đua lớp 8B3</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="sidebar-section-label">👤 Vai trò truy cập</div>', unsafe_allow_html=True)
    vai_tro = st.selectbox(
        "Chọn vai trò:",
        ["👑 Giáo viên chủ nhiệm", "Tổ trưởng Tổ 1", "Tổ trưởng Tổ 2", "Tổ trưởng Tổ 3", "Tổ trưởng Tổ 4"],
        label_visibility="collapsed",
    )

    st.markdown('<div class="sidebar-section-label">🧭 Bảng chức năng</div>', unsafe_allow_html=True)
    menu_items = [
        "📝 Ghi nhận điểm số",
        "✏️ Chỉnh sửa / Khắc phục",
        "📊 Tổng quan & Thống kê",
        "📥 Xuất dữ liệu báo cáo",
    ]
    if "Giáo viên" in vai_tro:
        menu_items.append("⚙️ Khởi tạo tuần mới (Reset)")

    chuc_nang = st.radio("Chọn chức năng:", menu_items, label_visibility="collapsed")

    st.markdown("<hr style='margin: 18px 0 14px 0; border: none; border-top: 1px solid #e2e8f0;'>", unsafe_allow_html=True)

    if "Giáo viên" in vai_tro:
        with st.expander("📁 Tải danh sách lớp (Excel/CSV)"):
            uploaded_file = st.file_uploader("Chọn file từ máy:", type=["xlsx", "csv"])
            if uploaded_file is not None:
                try:
                    if uploaded_file.name.endswith(".xlsx"):
                        new_df = pd.read_excel(uploaded_file)
                    else:
                        new_df = pd.read_csv(uploaded_file)

                    new_df = clean_dataframe(new_df)
                    st.session_state.df = new_df
                    new_df.to_csv(DATA_FILE, index=False)
                    st.success("✅ Đã chuẩn hóa danh sách thành công!")
                    time.sleep(1)
                    st.rerun()
                except Exception as e:
                    st.error(f"Lỗi đọc file: {e}")

    st.write("")
    if st.button("🔄 Đồng bộ lại dữ liệu", use_container_width=True):
        st.session_state.df = load_data()
        st.toast("Đã đồng bộ lại dữ liệu từ Google Sheets!", icon="🔄")
        time.sleep(0.5)
        st.rerun()

# ==============================================================================
# 6. PHÂN QUYỀN TRUY CẬP: TỔ NÀO CHỈ ĐƯỢC NHẬP VÀ SỬA TỔ ĐÓ
# ==============================================================================
is_gvcn = "Giáo viên" in vai_tro

if "Tổ 1" in vai_tro:
    df_role = df[df["Tổ"].astype(str).str.contains("1", na=False)].copy()
elif "Tổ 2" in vai_tro:
    df_role = df[df["Tổ"].astype(str).str.contains("2", na=False)].copy()
elif "Tổ 3" in vai_tro:
    df_role = df[df["Tổ"].astype(str).str.contains("3", na=False)].copy()
elif "Tổ 4" in vai_tro:
    df_role = df[df["Tổ"].astype(str).str.contains("4", na=False)].copy()
else:
    df_role = df.copy()

# ==============================================================================
# 7. BANNER TRƯỜNG TRÊN ĐẦU NỘI DUNG CHÍNH
# ==============================================================================
phan_vi_text = "Được thao tác: Toàn bộ lớp 8B3" if is_gvcn else f"Phạm vi giới hạn: {vai_tro}"
st.markdown(
    f"""
    <div class="school-top-bar">
        <div>
            <div class="school-title">TRƯỜNG THCS CHÍ CÔNG - PHAN RÍ CỬA</div>
            <div class="school-slogan">HỆ THỐNG THEO DÕI THI ĐUA & NỀ NẾP LỚP 8B3</div>
        </div>
        <div style="text-align: right; font-size: 0.85rem;">
            <div><b>Vai trò:</b> {vai_tro}</div>
            <div style="opacity: 0.85;">{phan_vi_text}</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ==============================================================================
# MÀN HÌNH 1: GHI NHẬN ĐIỂM SỐ
# ==============================================================================
if chuc_nang == "📝 Ghi nhận điểm số":
    st.subheader(f"📝 Ghi nhận điểm thi đua — {vai_tro}")
    st.caption("Cộng hoặc trừ điểm thi đua cho học sinh theo các mục đánh giá nề nếp")

    danh_sach_hien_thi = [
        f"#{row['STT']} - {row['Họ và tên']} ({row['Tổ']})"
        for _, row in df_role.iterrows()
    ]

    with st.container(border=True):
        col1, col2 = st.columns([1.6, 1])
        with col1:
            hs_chon_raw = st.selectbox(
                "👤 Chọn học sinh:",
                danh_sach_hien_thi if danh_sach_hien_thi else ["(Không có học sinh trong phạm vi phụ trách)"],
            )
            loi_chon = st.selectbox(
                "📋 Nội dung tuyên dương / Vi phạm:",
                list(DANH_SACH_LOI.keys()),
            )
        with col2:
            so_lan = st.number_input("🔢 Số lần thực hiện:", min_value=1, max_value=20, value=1, step=1)
            ghi_chu = st.text_input("📝 Ghi chú chi tiết (VD: Tiết 3 Hóa, BTVN Đại số):")

        # Xem trước biến động điểm
        if danh_sach_hien_thi and not hs_chon_raw.startswith("("):
            stt_hs = int(hs_chon_raw.split(" - ")[0].replace("#", ""))
            current_row = df[df["STT"] == stt_hs].iloc[0]
            diem_delta = DANH_SACH_LOI[loi_chon] * so_lan
            diem_sau_khi_luu = current_row["Điểm thi đua"] + diem_delta
            delta_color = "#16a34a" if diem_delta > 0 else "#dc2626"
            delta_sign = "+" if diem_delta > 0 else ""

            st.markdown(
                f"""
                <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; padding: 10px 16px; border-radius: 8px; margin: 12px 0;">
                    <b>Điểm hiện tại:</b> <span style="color:#0284c7; font-weight:700;">{current_row['Điểm thi đua']}</span> | 
                    <b>Biến động:</b> <span style="color:{delta_color}; font-weight:700;">{delta_sign}{diem_delta}</span> | 
                    <b>Dự kiến sau ghi:</b> <span style="color:#0f172a; font-weight:700;">{diem_sau_khi_luu}</span> ({tinh_xep_loai(diem_sau_khi_luu)})
                </div>
                """,
                unsafe_allow_html=True,
            )

        btn_save = st.button("💾 Xác nhận lưu điểm", type="primary", use_container_width=True)

        if btn_save and danh_sach_hien_thi and not hs_chon_raw.startswith("("):
            stt_hs = int(hs_chon_raw.split(" - ")[0].replace("#", ""))
            match_idx = df[df["STT"] == stt_hs].index[0]

            diem_thay_doi = DANH_SACH_LOI[loi_chon] * int(so_lan)
            old_score = int(df.at[match_idx, "Điểm thi đua"])
            new_score = old_score + diem_thay_doi
            xeploai_moi = tinh_xep_loai(new_score)

            noi_dung_ngan = loi_chon.split(" (")[0]
            log_item = f"{noi_dung_ngan} x{so_lan}" + (f" [{ghi_chu}]" if ghi_chu else "")

            old_log = str(df.at[match_idx, "Lỗi vi phạm"]).strip()
            if old_log in ["None", "nan", "", "NaN"]:
                full_log = log_item
            else:
                full_log = f"{old_log} | {log_item}"

            df.at[match_idx, "Điểm thi đua"] = new_score
            df.at[match_idx, "Xếp loại"] = xeploai_moi
            df.at[match_idx, "Lỗi vi phạm"] = full_log
            df.to_csv(DATA_FILE, index=False)
            st.session_state.df = df

            if API_URL and "/exec" in API_URL:
                ten_hs = df.at[match_idx, "Họ và tên"]
                ok, msg = update_gsheet(ten_hs, new_score, xeploai_moi, full_log)
                if not ok:
                    st.warning(f"Lưu máy thành công, Cloud báo: {msg}")

            st.toast(f"Đã lưu điểm cho {df.at[match_idx, 'Họ và tên']}!", icon="✅")
            time.sleep(0.8)
            st.rerun()

    st.markdown(f"#### 📋 Danh sách học sinh ({vai_tro})")
    st.dataframe(
        df_role[["STT", "Họ và tên", "Tổ", "Điểm thi đua", "Xếp loại", "Lỗi vi phạm"]],
        use_container_width=True,
        hide_index=True,
    )

# ==============================================================================
# MÀN HÌNH 2: CHỈNH SỬA / KHẮC PHỤC
# ==============================================================================
elif chuc_nang == "✏️ Chỉnh sửa / Khắc phục":
    st.subheader(f"✏️ Chỉnh sửa & Khắc phục sai sót — {vai_tro}")
    st.caption("Chỉ được sửa/hoàn tác các mục trong phạm vi học sinh thuộc tổ của mình.")

    danh_sach_hien_thi = [
        f"#{row['STT']} - {row['Họ và tên']} ({row['Tổ']})"
        for _, row in df_role.iterrows()
    ]

    if not danh_sach_hien_thi:
        st.warning("⚠️ Hiện không tìm thấy học sinh nào thuộc phân ban phụ trách của bạn!")
    else:
        selected_student_str = st.selectbox("👤 Chọn học sinh cần chỉnh sửa:", danh_sach_hien_thi)

        if selected_student_str:
            stt_target = int(selected_student_str.split(" - ")[0].replace("#", ""))
            match_idx = df[df["STT"] == stt_target].index[0]
            student_data = df.loc[match_idx]

            curr_score = int(student_data["Điểm thi đua"])
            curr_rank = student_data["Xếp loại"]
            raw_log = str(student_data["Lỗi vi phạm"]).strip()

            badge_bg = "#dcfce7" if curr_score >= 100 else ("#fef9c3" if curr_score >= 85 else "#fee2e2")
            badge_text = "#15803d" if curr_score >= 100 else ("#a16207" if curr_score >= 85 else "#b91c1c")
            score_color = "#15803d" if curr_score >= 100 else ("#b45309" if curr_score >= 85 else "#dc2626")
            letter_avatar = student_data['Họ và tên'].split()[-1][0].upper()

            st.markdown(
                f"""
                <div class="student-profile-card">
                    <div style="display: flex; align-items: center; gap: 14px;">
                        <div class="student-avatar">{letter_avatar}</div>
                        <div>
                            <div style="font-size: 1.15rem; font-weight: 700; color: #1e293b;">{student_data['Họ và tên']}</div>
                            <div style="font-size: 0.85rem; color: #64748b;">Thuộc phân ban: <b>{student_data['Tổ']}</b> | STT: #{student_data['STT']}</div>
                        </div>
                    </div>
                    <div style="display: flex; gap: 16px; align-items: center;">
                        <div style="text-align: right;">
                            <div style="font-size: 0.75rem; color: #64748b; text-transform: uppercase; font-weight: 600;">Điểm thi đua</div>
                            <div style="font-size: 1.5rem; font-weight: 800; color: {score_color};">{curr_score} đ</div>
                        </div>
                        <span style="background-color: {badge_bg}; color: {badge_text}; font-weight: 700; padding: 6px 14px; border-radius: 30px; font-size: 0.85rem; border: 1px solid {badge_text}30;">
                            Xếp loại: {curr_rank}
                        </span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            st.markdown("#### 📋 Danh sách các mục đã ghi nhận trong tuần:")

            if raw_log in ["None", "nan", "", "NaN"]:
                items_list = []
            else:
                items_list = [item.strip() for item in raw_log.split("|") if item.strip()]

            if not items_list:
                st.markdown(
                    """
                    <div class="empty-log-box">
                        <div style="font-size: 2.2rem; margin-bottom: 6px;">🌟</div>
                        <div style="font-weight: 700; color: #0b4f8a; font-size: 1.05rem;">Học sinh này hiện chưa có vi phạm hoặc điểm trừ nào!</div>
                        <div style="color: #64748b; font-size: 0.88rem;">Nề nếp và tinh thần học tập đang được duy trì rất tốt.</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            else:
                for idx, item in enumerate(items_list):
                    matched_loi, so_lan_old, ghi_chu_old, delta_old = parse_item_info(item)
                    color_score = "#16a34a" if delta_old > 0 else "#dc2626"
                    sign = "+" if delta_old > 0 else ""

                    with st.container(border=True):
                        col_info, col_act1, col_act2 = st.columns([3.5, 1, 1])
                        with col_info:
                            st.markdown(
                                f"**#{idx + 1}.** `{item}` &nbsp; <b style='color:{color_score};'>({sign}{delta_old} điểm)</b>",
                                unsafe_allow_html=True,
                            )
                        with col_act1:
                            edit_key = f"is_editing_{stt_target}_{idx}"
                            if edit_key not in st.session_state:
                                st.session_state[edit_key] = False

                            if st.button("✏️ Sửa mục này", key=f"btn_toggle_{stt_target}_{idx}", use_container_width=True):
                                st.session_state[edit_key] = not st.session_state[edit_key]
                                st.rerun()

                        with col_act2:
                            if st.button("🗑️ Xóa bỏ", key=f"del_item_{stt_target}_{idx}", use_container_width=True):
                                new_score = curr_score - delta_old
                                new_rank = tinh_xep_loai(new_score)

                                items_list.pop(idx)
                                new_log = " | ".join(items_list) if items_list else ""

                                df.at[match_idx, "Điểm thi đua"] = new_score
                                df.at[match_idx, "Xếp loại"] = new_rank
                                df.at[match_idx, "Lỗi vi phạm"] = new_log
                                df.to_csv(DATA_FILE, index=False)
                                st.session_state.df = df

                                if API_URL and "/exec" in API_URL:
                                    update_gsheet(student_data["Họ và tên"], new_score, new_rank, new_log)

                                st.toast("Đã xóa mục và hoàn lại điểm!", icon="✅")
                                time.sleep(0.8)
                                st.rerun()

                        if st.session_state.get(f"is_editing_{stt_target}_{idx}", False):
                            st.markdown("---")
                            st.caption(f"🔧 Đang chỉnh sửa mục số {idx + 1}:")
                            e_col1, e_col2 = st.columns([2, 1])

                            all_keys = list(DANH_SACH_LOI.keys())
                            default_index = all_keys.index(matched_loi) if matched_loi in all_keys else 0

                            with e_col1:
                                new_loi_selected = st.selectbox(
                                    "Chọn lại nội dung đúng:",
                                    all_keys,
                                    index=default_index,
                                    key=f"select_loi_{stt_target}_{idx}",
                                )
                                new_ghichu_edit = st.text_input(
                                    "Ghi chú sửa đổi:",
                                    value=ghi_chu_old,
                                    key=f"ghichu_edit_{stt_target}_{idx}",
                                )
                            with e_col2:
                                new_solan_edit = st.number_input(
                                    "Số lần đúng:",
                                    min_value=1,
                                    max_value=20,
                                    value=int(so_lan_old),
                                    key=f"solan_edit_{stt_target}_{idx}",
                                )
                                st.write("")
                                if st.button("💾 Lưu cập nhật", type="primary", key=f"save_edit_{stt_target}_{idx}", use_container_width=True):
                                    new_delta = DANH_SACH_LOI[new_loi_selected] * int(new_solan_edit)
                                    score_diff = new_delta - delta_old
                                    new_score = curr_score + score_diff
                                    new_rank = tinh_xep_loai(new_score)

                                    ten_rut_gon = new_loi_selected.split(" (")[0]
                                    new_item_str = f"{ten_rut_gon} x{new_solan_edit}" + (f" [{new_ghichu_edit}]" if new_ghichu_edit else "")
                                    items_list[idx] = new_item_str
                                    new_log = " | ".join(items_list)

                                    df.at[match_idx, "Điểm thi đua"] = new_score
                                    df.at[match_idx, "Xếp loại"] = new_rank
                                    df.at[match_idx, "Lỗi vi phạm"] = new_log
                                    df.to_csv(DATA_FILE, index=False)
                                    st.session_state.df = df

                                    if API_URL and "/exec" in API_URL:
                                        update_gsheet(student_data["Họ và tên"], new_score, new_rank, new_log)

                                    st.session_state[f"is_editing_{stt_target}_{idx}"] = False
                                    st.success("Đã cập nhật mục sửa đổi thành công!")
                                    time.sleep(0.8)
                                    st.rerun()

            st.markdown("---")

            with st.expander("➕ Ghi nhận thêm điểm thưởng / vi phạm ngay tại đây"):
                a_col1, a_col2 = st.columns([2, 1])
                with a_col1:
                    add_loi = st.selectbox("Nội dung cần thêm:", list(DANH_SACH_LOI.keys()), key=f"add_loi_{stt_target}")
                    add_note = st.text_input("Ghi chú kèm theo:", key=f"add_note_{stt_target}")
                with a_col2:
                    add_times = st.number_input("Số lần:", min_value=1, max_value=20, value=1, key=f"add_times_{stt_target}")
                    st.write("")
                    if st.button("➕ Thêm vào hồ sơ", type="primary", key=f"btn_add_now_{stt_target}", use_container_width=True):
                        add_delta = DANH_SACH_LOI[add_loi] * int(add_times)
                        new_score = curr_score + add_delta
                        new_rank = tinh_xep_loai(new_score)

                        item_add_str = f"{add_loi.split(' (')[0]} x{add_times}" + (f" [{add_note}]" if add_note else "")
                        if raw_log in ["None", "nan", "", "NaN"]:
                            new_log = item_add_str
                        else:
                            new_log = f"{raw_log} | {item_add_str}"

                        df.at[match_idx, "Điểm thi đua"] = new_score
                        df.at[match_idx, "Xếp loại"] = new_rank
                        df.at[match_idx, "Lỗi vi phạm"] = new_log
                        df.to_csv(DATA_FILE, index=False)
                        st.session_state.df = df

                        if API_URL and "/exec" in API_URL:
                            update_gsheet(student_data["Họ và tên"], new_score, new_rank, new_log)

                        st.toast("Đã thêm mục mới thành công!", icon="✅")
                        time.sleep(0.8)
                        st.rerun()

            if is_gvcn:
                with st.expander("🛠️ Can thiệp sửa điểm thủ công (Đặc quyền GVCN)"):
                    c_s1, c_s2 = st.columns(2)
                    with c_s1:
                        manual_score = st.number_input("Điểm muốn đặt lại:", value=curr_score, step=1)
                    with c_s2:
                        st.write("")
                        st.write("")
                        btn_manual = st.button("Lưu điểm tự nhập")

                    if btn_manual:
                        new_manual_rank = tinh_xep_loai(manual_score)
                        df.at[match_idx, "Điểm thi đua"] = int(manual_score)
                        df.at[match_idx, "Xếp loại"] = new_manual_rank
                        df.to_csv(DATA_FILE, index=False)
                        st.session_state.df = df

                        if API_URL and "/exec" in API_URL:
                            update_gsheet(student_data["Họ và tên"], int(manual_score), new_manual_rank, raw_log)

                        st.success("Đã cập nhật lại điểm thủ công!")
                        time.sleep(0.8)
                        st.rerun()

# ==============================================================================
# MÀN HÌNH 3: TỔNG QUAN & THỐNG KÊ
# ==============================================================================
elif chuc_nang == "📊 Tổng quan & Thống kê":
    st.subheader("📊 Bảng thống kê toàn diện lớp học")

    m1, m2, m3, m4 = st.columns(4)
    total_students = len(df)
    avg_score = round(df["Điểm thi đua"].mean(), 1) if not df.empty else 0
    tot_count = len(df[df["Xếp loại"] == "Tốt"])
    yeu_count = len(df[df["Xếp loại"].isin(["Trung bình", "Yếu"])])

    with m1:
        st.markdown(
            f"""
            <div class="stat-card">
                <span class="stat-label">Sĩ số lớp</span>
                <span class="stat-value">{total_students}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with m2:
        st.markdown(
            f"""
            <div class="stat-card">
                <span class="stat-label">Điểm TB Lớp</span>
                <span class="stat-value" style="color:#0284c7;">{avg_score}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with m3:
        st.markdown(
            f"""
            <div class="stat-card">
                <span class="stat-label">Xếp loại Tốt</span>
                <span class="stat-value" style="color:#16a34a;">{tot_count}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with m4:
        st.markdown(
            f"""
            <div class="stat-card">
                <span class="stat-label">Cần nhắc nhở</span>
                <span class="stat-value" style="color:#dc2626;">{yeu_count}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("---")

    if "Tổ" in df.columns and not df.empty:
        st.markdown("#### 🏆 Điểm trung bình thi đua theo Tổ")
        avg_by_to = df.groupby("Tổ")["Điểm thi đua"].mean().round(1).reset_index()
        avg_by_to = avg_by_to.sort_values(by="Điểm thi đua", ascending=False)
        st.bar_chart(data=avg_by_to.set_index("Tổ"), use_container_width=True)

    st.markdown("#### 📄 Bảng tổng hợp thi đua")
    f_col1, f_col2 = st.columns(2)
    with f_col1:
        loc_to = st.selectbox("Lọc theo tổ:", ["Tất cả"] + sorted(list(df["Tổ"].unique())))
    with f_col2:
        loc_xeploai = st.selectbox("Lọc theo loại:", ["Tất cả"] + ["Tốt", "Khá", "Trung bình", "Yếu"])

    df_view = df.copy()
    if loc_to != "Tất cả":
        df_view = df_view[df_view["Tổ"] == loc_to]
    if loc_xeploai != "Tất cả":
        df_view = df_view[df_view["Xếp loại"] == loc_xeploai]

    st.dataframe(
        df_view.sort_values(by="Điểm thi đua", ascending=False),
        column_config={
            "Điểm thi đua": st.column_config.ProgressColumn(
                "Điểm thi đua",
                help="Thang điểm chuẩn 100",
                format="%d",
                min_value=0,
                max_value=120,
            ),
        },
        use_container_width=True,
        hide_index=True,
    )

# ==============================================================================
# MÀN HÌNH 4: XUẤT BÁO CÁO EXCEL ĐÃ SẮP XẾP ĐIỂM TỪ CAO XUỐNG THẤP
# ==============================================================================
elif chuc_nang == "📥 Xuất dữ liệu báo cáo":
    st.subheader("📥 Xuất báo cáo tổng kết thi đua")
    st.caption("File Excel được tạo tự động: xếp hạng học sinh có điểm cao từ trên xuống dưới.")

    with st.container(border=True):
        st.write("📌 Danh sách trong file Excel đã được **sắp xếp theo điểm thi đua giảm dần (Top điểm cao ở trên cùng)** và đánh lại số thứ tự xếp hạng rõ ràng.")

        def tao_excel_chuyen_nghiep(dataframe):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Báo Cáo Thi Đua"
            ws.views.sheetView[0].showGridLines = True

            font_title = Font(name="Calibri", size=15, bold=True, color="0B4F8A")
            font_sub = Font(name="Calibri", size=11, italic=True, color="475569")
            font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            font_data = Font(name="Calibri", size=11)
            font_bold = Font(name="Calibri", size=11, bold=True)

            fill_header = PatternFill(start_color="0B4F8A", end_color="0B4F8A", fill_type="solid")
            fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
            fill_total = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

            border_thin = Side(border_style="thin", color="CBD5E1")
            border_full = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)

            # 1. Tiêu đề trường & bảng
            ws.merge_cells("A1:F1")
            ws["A1"] = "TRƯỜNG THCS CHÍ CÔNG - PHAN RÍ CỬA"
            ws["A1"].font = font_title
            ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

            ws.merge_cells("A2:F2")
            ws["A2"] = f"BẢNG THỐNG KÊ THI ĐUA LỚP 8B3 (Xuất ngày: {time.strftime('%d/%m/%Y %H:%M')})"
            ws["A2"].font = font_sub
            ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

            ws.append([])

            # 2. Tiêu đề cột
            headers = ["Hạng", "Họ và tên", "Tổ", "Điểm thi đua", "Xếp loại", "Chi tiết vi phạm / Tuyên dương"]
            ws.append(headers)
            header_row = 4

            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=header_row, column=col_idx)
                cell.font = font_header
                cell.fill = fill_header
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = border_full

            ws.row_dimensions[header_row].height = 26

            # 3. Sắp xếp dữ liệu từ điểm cao nhất xuống thấp nhất
            df_sorted = dataframe.sort_values(by=["Điểm thi đua", "Họ và tên"], ascending=[False, True]).reset_index(drop=True)

            current_row = 5
            for idx, row in df_sorted.iterrows():
                raw_val = str(row.get("Lỗi vi phạm", "")).strip()
                score_val = int(row.get("Điểm thi đua", 100))
                
                # Nếu chuỗi trống mà điểm khác 100, tự động ghi chú biến động điểm
                if raw_val in ["None", "nan", "NaN", ""]:
                    if score_val > 100:
                        noi_dung_xuat = f"Tuyên dương (+{score_val - 100} điểm)"
                    elif score_val < 100:
                        noi_dung_xuat = f"Vi phạm nề nếp ({score_val - 100} điểm)"
                    else:
                        noi_dung_xuat = "—"
                else:
                    noi_dung_xuat = raw_val

                # Ghi số thứ tự xếp hạng (1, 2, 3...) theo điểm cao xuống thấp
                ws.append([
                    idx + 1,
                    row.get("Họ và tên", ""),
                    row.get("Tổ", ""),
                    score_val,
                    row.get("Xếp loại", "Tốt"),
                    noi_dung_xuat
                ])

                for c_idx in range(1, 7):
                    cell = ws.cell(row=current_row, column=c_idx)
                    cell.font = font_data
                    cell.border = border_full
                    if current_row % 2 == 0:
                        cell.fill = fill_zebra

                    if c_idx in [1, 3, 5]:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    elif c_idx == 4:
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    elif c_idx == 6:
                        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")

                ws.row_dimensions[current_row].height = 24
                current_row += 1

            # 4. Hàng tổng kết điểm trung bình
            ws.merge_cells(f"A{current_row}:C{current_row}")
            total_label = ws.cell(row=current_row, column=1, value="ĐIỂM TRUNG BÌNH CẢ LỚP")
            total_label.font = font_bold
            total_label.alignment = Alignment(horizontal="center", vertical="center")

            avg_val = round(dataframe["Điểm thi đua"].mean(), 1) if not dataframe.empty else 0
            avg_cell = ws.cell(row=current_row, column=4, value=avg_val)
            avg_cell.font = font_bold
            avg_cell.alignment = Alignment(horizontal="right", vertical="center")

            for c_idx in range(1, 7):
                c = ws.cell(row=current_row, column=c_idx)
                c.fill = fill_total
                c.border = border_full

            ws.row_dimensions[current_row].height = 24

            # 5. Căn chỉnh độ rộng cột
            col_widths = {1: 8, 2: 26, 3: 12, 4: 15, 5: 14, 6: 60}
            for col_idx, width in col_widths.items():
                col_letter = get_column_letter(col_idx)
                ws.column_dimensions[col_letter].width = width

            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            return output

        df_xuat = st.session_state.df.copy()
        excel_data = tao_excel_chuyen_nghiep(df_xuat)

        st.write("")
        st.download_button(
            label="📊 Tải xuống file Excel Báo Cáo (.xlsx)",
            data=excel_data,
            file_name=f"Bao_Cao_Thi_Dua_Lop_8B3_{time.strftime('%Y_%m_%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

# ==============================================================================
# MÀN HÌNH 5: RESET TUẦN MỚI
# ==============================================================================
elif chuc_nang == "⚙️ Khởi tạo tuần mới (Reset)":
    st.subheader("⚙️ Khởi tạo chu kỳ tuần mới")
    st.warning("⚠️ **Lưu ý:** Hãy tải file báo cáo Excel ở mục '📥 Xuất dữ liệu báo cáo' trước khi thực hiện thao tác này.")

    with st.container(border=True):
        st.markdown(
            """
            **Hệ thống sẽ thực hiện:**
            - Khôi phục điểm tất cả học sinh về **100 điểm**.
            - Cập nhật toàn bộ xếp loại về **Tốt**.
            - Xóa sạch danh sách ghi nhận lỗi vi phạm/khen thưởng cũ.
            """
        )

        confirm = st.checkbox("Xác nhận: Tôi đã lưu trữ báo cáo tuần này và muốn đặt lại dữ liệu.")

        if st.button("🔄 Bắt đầu tuần mới", type="primary"):
            if confirm:
                with st.spinner("Đang đặt lại dữ liệu lớp..."):
                    df["Điểm thi đua"] = 100
                    df["Xếp loại"] = "Tốt"
                    df["Lỗi vi phạm"] = ""
                    df.to_csv(DATA_FILE, index=False)
                    st.session_state.df = df

                    if API_URL and "/exec" in API_URL:
                        ok, msg = reset_new_week_gsheet()
                        if not ok:
                            st.warning(f"Đã lưu local, cảnh báo kết nối Cloud: {msg}")
                        else:
                            st.success("Đã đồng bộ lên Google Sheets thành công!")

                    st.toast("Đã khởi tạo tuần mới!", icon="🎉")
                    time.sleep(1)
                    st.rerun()
            else:
                st.error("Vui lòng tích vào ô xác nhận trước khi tiếp tục!")
