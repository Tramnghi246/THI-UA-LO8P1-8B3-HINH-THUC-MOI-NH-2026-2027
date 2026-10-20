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
# 1. CẤU HÌNH TRANG MOBILE-FIRST (ẨN HOÀN TOÀN SIDEBAR)
# ==============================================================================
st.set_page_config(
    page_title="App Thi Đua 8B3",
    page_icon="📱",
    layout="centered", # Gọn gàng như màn hình điện thoại
    initial_sidebar_state="collapsed",
)

API_URL = "https://script.google.com/macros/s/AKfycbwoy6GT8XEkwwrLTSXrIk3tMj6lzATeiwrw02cWYz7yEXDhzEtDOoyL8VssgXuJE08q/exec"
DATA_FILE = "du_lieu_thi_dua.csv"

# ==============================================================================
# 2. CSS BIẾN WEB THÀNH APP MOBILE CHUYÊN NGHIỆP
# ==============================================================================
st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Segoe+UI:wght@400;600;700&display=swap');
        html, body, [class*="css"] {
            font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
        }

        /* Ẩn hoàn toàn thanh Sidebar và nút mở Sidebar mặc định */
        section[data-testid="stSidebar"], [data-testid="collapsedControl"] {
            display: none !important;
        }

        /* Chừa khoảng trống dưới đáy cho thanh menu cố định */
        .block-container {
            padding-top: 1rem !important;
            padding-bottom: 5.5rem !important;
            max-width: 650px !important; /* Độ rộng chuẩn màn hình Mobile / Tablet */
        }

        /* Header App Mobile màu xanh gradient */
        .mobile-header {
            background: linear-gradient(135deg, #0b4f8a 0%, #1976d2 100%);
            color: #ffffff;
            padding: 16px 20px;
            border-radius: 16px;
            margin-bottom: 16px;
            box-shadow: 0 4px 14px rgba(11, 79, 138, 0.25);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .mobile-header-title {
            font-size: 1.2rem;
            font-weight: 800;
            letter-spacing: 0.5px;
        }
        .mobile-header-sub {
            font-size: 0.8rem;
            color: #ffeb3b;
            font-weight: 600;
        }

        /* Thẻ học sinh bo cong dạng Card Mobile */
        .student-mobile-card {
            background: #ffffff;
            border-radius: 14px;
            padding: 14px 18px;
            margin-bottom: 12px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 2px 6px rgba(0,0,0,0.03);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .student-mobile-avatar {
            width: 44px;
            height: 44px;
            background: #e0f2fe;
            color: #0284c7;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: 800;
            font-size: 1.15rem;
            margin-right: 12px;
        }

        /* Thẻ thống kê 2 cột */
        .stat-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 12px;
            margin-bottom: 16px;
        }
        .stat-card-m {
            background: #ffffff;
            border-radius: 14px;
            padding: 14px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 2px 6px rgba(0,0,0,0.03);
            text-align: center;
        }
        .stat-card-m-title {
            font-size: 0.75rem;
            color: #64748b;
            text-transform: uppercase;
            font-weight: 700;
        }
        .stat-card-m-val {
            font-size: 1.6rem;
            font-weight: 800;
            color: #0f172a;
            margin-top: 4px;
        }

        /* Nút bấm bo tròn mượt mà */
        .stButton>button {
            border-radius: 10px;
            font-weight: 700;
            padding: 8px 16px;
            transition: all 0.2s ease;
        }

        /* Thanh điều hướng cố định dưới đáy màn hình (Bottom Navigation Bar) */
        .bottom-nav-container {
            position: fixed;
            bottom: 0;
            left: 0;
            right: 0;
            background: #ffffff;
            border-top: 1px solid #e2e8f0;
            padding: 8px 12px;
            display: flex;
            justify-content: space-around;
            align-items: center;
            box-shadow: 0 -2px 10px rgba(0,0,0,0.05);
            z-index: 999999;
            max-width: 650px;
            margin: 0 auto;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==============================================================================
# 3. DANH MỤC THI ĐUA & DỮ LIỆU
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

# Khởi tạo tab điều hướng mobile
if "current_tab" not in st.session_state:
    st.session_state.current_tab = "Theo dõi"

# ==============================================================================
# 5. HEADER MOBILE & THANH CHỌN VAI TRÒ
# ==============================================================================
st.markdown(
    """
    <div class="mobile-header">
        <div>
            <div class="mobile-header-title">📱 LỚP 8B3 - CHÍ CÔNG</div>
            <div class="mobile-header-sub">HỆ THỐNG THI ĐUA HỌC ĐƯỜNG</div>
        </div>
        <div style="font-size: 1.6rem;">🏫</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# Thanh chọn vai trò nhanh đặt ngay trên màn hình chính
top_col1, top_col2 = st.columns([1.8, 1])
with top_col1:
    vai_tro = st.selectbox(
        "Vai trò đang làm việc:",
        ["👑 Giáo viên chủ nhiệm", "Tổ trưởng Tổ 1", "Tổ trưởng Tổ 2", "Tổ trưởng Tổ 3", "Tổ trưởng Tổ 4"],
        label_visibility="collapsed",
    )
with top_col2:
    if st.button("🔄 Đồng bộ", use_container_width=True):
        st.session_state.df = load_data()
        st.toast("Đã đồng bộ dữ liệu mới nhất!", icon="🔄")
        time.sleep(0.5)
        st.rerun()

is_gvcn = "Giáo viên" in vai_tro

# Lọc theo vai trò
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

st.markdown("<hr style='margin: 10px 0 16px 0; border: none; border-top: 1px solid #e2e8f0;'>", unsafe_allow_html=True)

# ==============================================================================
# 6. NỘI DUNG 4 MÀN HÌNH CHÍNH THEO TAB MOBILE
# ==============================================================================

# --- TAB 1: THEO DÕI DANH SÁCH DẠNG CARD CHẠM ---
if st.session_state.current_tab == "Theo dõi":
    st.markdown("##### 📋 Danh Sách Điểm Thi Đua")
    
    # Bộ lọc nhanh theo tổ nếu là GVCN
    if is_gvcn:
        selected_group = st.pills("Phân ban:", ["Tất cả", "Tổ 1", "Tổ 2", "Tổ 3", "Tổ 4"], default="Tất cả")
        df_show = df_role.copy() if selected_group == "Tất cả" else df_role[df_role["Tổ"] == selected_group]
    else:
        df_show = df_role.copy()

    # Hiển thị từng học sinh dạng Thẻ bo tròn mềm mại
    for idx, row in df_show.iterrows():
        name = row["Họ và tên"]
        score = row["Điểm thi đua"]
        group = row["Tổ"]
        rank = row["Xếp loại"]
        first_letter = name.split()[-1][0].upper()
        
        score_color = "#16a34a" if score >= 100 else ("#d97706" if score >= 85 else "#dc2626")
        badge_bg = "#dcfce7" if score >= 100 else ("#fef3c7" if score >= 85 else "#fee2e2")

        st.markdown(
            f"""
            <div class="student-mobile-card">
                <div style="display: flex; align-items: center;">
                    <div class="student-mobile-avatar">{first_letter}</div>
                    <div>
                        <div style="font-size: 1rem; font-weight: 700; color: #1e293b;">{name}</div>
                        <div style="font-size: 0.78rem; color: #64748b;">{group} • STT #{row['STT']}</div>
                    </div>
                </div>
                <div style="text-align: right;">
                    <div style="font-size: 1.25rem; font-weight: 800; color: {score_color};">{score} đ</div>
                    <span style="background-color: {badge_bg}; color: {score_color}; font-size: 0.72rem; font-weight: 700; padding: 2px 8px; border-radius: 20px;">
                        {rank}
                    </span>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

# --- TAB 2: CHẤM ĐIỂM & VI PHẠM ---
elif st.session_state.current_tab == "Chấm điểm":
    st.markdown(f"##### ✍️ Ghi Nhận Thi Đua ({vai_tro})")

    danh_sach_hs = [
        f"#{row['STT']} - {row['Họ và tên']} ({row['Tổ']})"
        for _, row in df_role.iterrows()
    ]

    with st.container(border=True):
        hs_chon_raw = st.selectbox("👤 Chọn học sinh:", danh_sach_hs if danh_sach_hs else ["(Không có học sinh)"])
        loi_chon = st.selectbox("📋 Nội dung tuyên dương / Vi phạm:", list(DANH_SACH_LOI.keys()))
        
        c1, c2 = st.columns([1, 1.5])
        with c1:
            so_lan = st.number_input("Số lần:", min_value=1, max_value=20, value=1)
        with c2:
            ghi_chu = st.text_input("Ghi chú chi tiết:")

        if danh_sach_hs and not hs_chon_raw.startswith("("):
            stt_hs = int(hs_chon_raw.split(" - ")[0].replace("#", ""))
            current_row = df[df["STT"] == stt_hs].iloc[0]
            diem_delta = DANH_SACH_LOI[loi_chon] * so_lan
            diem_sau_khi_luu = current_row["Điểm thi đua"] + diem_delta
            delta_color = "#16a34a" if diem_delta > 0 else "#dc2626"
            delta_sign = "+" if diem_delta > 0 else ""

            st.markdown(
                f"""
                <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; padding: 8px 12px; border-radius: 8px; margin: 8px 0; font-size: 0.88rem;">
                    Hiện tại: <b>{current_row['Điểm thi đua']}</b> | 
                    Biến động: <b style="color:{delta_color};">{delta_sign}{diem_delta}</b> | 
                    Dự kiến: <b>{diem_sau_khi_luu} đ</b>
                </div>
                """,
                unsafe_allow_html=True,
            )

        if st.button("💾 Xác Nhận Lưu Điểm", type="primary", use_container_width=True):
            if danh_sach_hs and not hs_chon_raw.startswith("("):
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
                    update_gsheet(df.at[match_idx, "Họ và tên"], new_score, xeploai_moi, full_log)

                st.toast(f"Đã lưu điểm cho {df.at[match_idx, 'Họ và tên']}!", icon="✅")
                time.sleep(0.5)
                st.rerun()

# --- TAB 3: THỐNG KÊ BIỂU ĐỒ MOBILE ---
elif st.session_state.current_tab == "Thống kê":
    st.markdown("##### 📊 Thống Kê Tổng Quan")

    total_students = len(df)
    avg_score = round(df["Điểm thi đua"].mean(), 1) if not df.empty else 0
    tot_count = len(df[df["Xếp loại"] == "Tốt"])
    yeu_count = len(df[df["Xếp loại"].isin(["Trung bình", "Yếu"])])

    st.markdown(
        f"""
        <div class="stat-grid">
            <div class="stat-card-m">
                <div class="stat-card-m-title">Sĩ số lớp</div>
                <div class="stat-card-m-val">{total_students}</div>
            </div>
            <div class="stat-card-m">
                <div class="stat-card-m-title">Điểm TB Lớp</div>
                <div class="stat-card-m-val" style="color:#0284c7;">{avg_score}</div>
            </div>
            <div class="stat-card-m">
                <div class="stat-card-m-title">Xếp loại Tốt</div>
                <div class="stat-card-m-val" style="color:#16a34a;">{tot_count}</div>
            </div>
            <div class="stat-card-m">
                <div class="stat-card-m-title">Nhắc nhở</div>
                <div class="stat-card-m-val" style="color:#dc2626;">{yeu_count}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if "Tổ" in df.columns and not df.empty:
        st.markdown("###### 🏆 Điểm trung bình 4 Tổ")
        avg_by_to = df.groupby("Tổ")["Điểm thi đua"].mean().round(1).reset_index()
        avg_by_to = avg_by_to.sort_values(by="Điểm thi đua", ascending=False)
        st.bar_chart(data=avg_by_to.set_index("Tổ"), use_container_width=True)

# --- TAB 4: BÁO CÁO & XUẤT EXCEL ---
elif st.session_state.current_tab == "Báo cáo":
    st.markdown("##### 📥 Xuất Báo Cáo Thi Đua")
    st.caption("File Excel tự động xếp hạng học sinh từ điểm cao nhất xuống thấp nhất.")

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

        ws.merge_cells("A1:F1")
        ws["A1"] = "TRƯỜNG THCS CHÍ CÔNG - PHAN RÍ CỬA"
        ws["A1"].font = font_title
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("A2:F2")
        ws["A2"] = f"BẢNG THỐNG KÊ THI ĐUA LỚP 8B3 (Xuất ngày: {time.strftime('%d/%m/%Y %H:%M')})"
        ws["A2"].font = font_sub
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

        ws.append([])

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

        df_sorted = dataframe.sort_values(by=["Điểm thi đua", "Họ và tên"], ascending=[False, True]).reset_index(drop=True)

        current_row = 5
        for idx, row in df_sorted.iterrows():
            raw_val = str(row.get("Lỗi vi phạm", "")).strip()
            score_val = int(row.get("Điểm thi đua", 100))
            
            if raw_val in ["None", "nan", "NaN", ""]:
                if score_val > 100:
                    noi_dung_xuat = f"Tuyên dương (+{score_val - 100} điểm)"
                elif score_val < 100:
                    noi_dung_xuat = f"Vi phạm nề nếp ({score_val - 100} điểm)"
                else:
                    noi_dung_xuat = "—"
            else:
                noi_dung_xuat = raw_val

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

        col_widths = {1: 8, 2: 26, 3: 12, 4: 15, 5: 14, 6: 60}
        for col_idx, width in col_widths.items():
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    excel_data = tao_excel_chuyen_nghiep(df)

    st.download_button(
        label="📊 Tải Xuống File Excel Báo Cáo",
        data=excel_data,
        file_name=f"Bao_Cao_Thi_Dua_Lop_8B3_{time.strftime('%Y_%m_%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    if is_gvcn:
        st.markdown("<hr style='margin: 20px 0;'>", unsafe_allow_html=True)
        st.markdown("###### ⚙️ Khởi Tạo Tuần Mới (Reset)")
        confirm = st.checkbox("Xác nhận đặt lại toàn bộ điểm về 100")
        if st.button("🔄 Bắt đầu tuần mới", type="primary", use_container_width=True):
            if confirm:
                df["Điểm thi đua"] = 100
                df["Xếp loại"] = "Tốt"
                df["Lỗi vi phạm"] = ""
                df.to_csv(DATA_FILE, index=False)
                st.session_state.df = df
                if API_URL and "/exec" in API_URL:
                    reset_new_week_gsheet()
                st.toast("Đã khởi tạo tuần mới!", icon="🎉")
                time.sleep(0.5)
                st.rerun()

# ==============================================================================
# 7. THANH BOTTOM NAVIGATION BAR (CỐ ĐỊNH DƯỚI ĐÁY MÀN HÌNH ĐIỆN THOẠI)
# ==============================================================================
st.markdown("<div style='height: 40px;'></div>", unsafe_allow_html=True)

nav_c1, nav_c2, nav_c3, nav_c4 = st.columns(4)

with nav_c1:
    if st.button("📋\nTheo dõi", key="btn_nav_1", use_container_width=True):
        st.session_state.current_tab = "Theo dõi"
        st.rerun()

with nav_c2:
    if st.button("✍️\nChấm điểm", key="btn_nav_2", use_container_width=True):
        st.session_state.current_tab = "Chấm điểm"
        st.rerun()

with nav_c3:
    if st.button("📊\nThống kê", key="btn_nav_3", use_container_width=True):
        st.session_state.current_tab = "Thống kê"
        st.rerun()

with nav_c4:
    if st.button("📥\nBáo cáo", key="btn_nav_4", use_container_width=True):
        st.session_state.current_tab = "Báo cáo"
        st.rerun()
