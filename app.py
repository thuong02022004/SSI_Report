import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime

# --- 1. KẾT NỐI GOOGLE SPREADSHEET ---
def get_google_spreadsheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds_dict = dict(st.secrets["google_creds"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    
    spreadsheet = client.open("Bảng tính không có tiêu đề")
    return spreadsheet

try:
    spreadsheet = get_google_spreadsheet()
except Exception as e:
    st.error(f"Lỗi kết nối Google Sheet: {e}")
    st.stop()

# --- 2. KHỞI TẠO SHEET GỐC ---
MAIN_SHEET_NAME = "Thông tin nhân viên"
all_worksheets = [ws.title for ws in spreadsheet.worksheets()]
main_sheet_obj = None

if MAIN_SHEET_NAME not in all_worksheets:
    try:
        new_main_sheet = spreadsheet.add_worksheet(title=MAIN_SHEET_NAME, rows="1000", cols="20")
        new_main_sheet.append_row(["ID", "Họ và tên", "Ngày sinh", "Tuổi", "Ngày bắt đầu", "Trạng thái", "Ghi chú"])
        all_worksheets.insert(0, MAIN_SHEET_NAME)
    except Exception as e:
        st.error(f"Không thể khởi tạo sheet gốc: {e}")

main_sheet_obj = spreadsheet.worksheet(MAIN_SHEET_NAME)


# --- 3. DANH SÁCH NGÀY LỄ VIỆT NAM & HÀM KIỂM TRA PHIÊN ---
def is_vietnam_holiday(dt):
    """Kiểm tra ngày được chọn có phải ngày lễ cố định hoặc ngày nghỉ của VN không"""
    md = dt.strftime("%m-%d")
    holidays = [
        "01-01",  # Tết Dương Lịch
        "04-30",  # Giải phóng miền Nam
        "05-01",  # Quốc tế Lao động
        "09-02",  # Quốc khánh
        "09-03"   # Quốc khánh (ngày bổ sung)
    ]
    lunar_holidays_2026 = ["02-16", "02-17", "02-18", "02-19", "02-20"]
    return (md in holidays) or (dt.strftime("%Y-%m-%d") in lunar_holidays_2026)

def calculate_trading_fees(trading_value):
    """
    HÀM TỰ ĐỘNG TÍNH PHÍ GIAO DỊCH VÀ PHÍ NET
    - Phí giao dịch công ty được hưởng = 0.25% giá trị giao dịch
    - Phí net của CTV = 30% của tiền phí giao dịch đó
    """
    company_fee = trading_value * 0.0025  # Công ty được 0.25%
    net_fee = company_fee * 0.30          # CTV được 30% của tiền phí giao dịch
    return company_fee, net_fee


# --- 4. CÁC HÀM XỬ LÝ DỮ LIỆU NHÂN VIÊN (CRUD) ---
def insert_employee(name, dob_str, age, start_date_str, role, note):
    records = main_sheet_obj.get_all_records()
    next_id_num = len(records) + 1
    emp_id = f"NV{next_id_num:03d}"
    
    new_row = [emp_id, name, dob_str, age, start_date_str, role, note]
    main_sheet_obj.append_row(new_row)
    
    current_sheets = [ws.title for ws in spreadsheet.worksheets()]
    if name not in current_sheets:
        employee_sheet = spreadsheet.add_worksheet(title=name, rows="1000", cols="20")
        employee_sheet.append_row(['Phiên giao dịch', 'Giá trị giao dịch', 'Phí giao dịch', 'Phí net', 'KH mới', 'KH chuyển ID'])
    return emp_id

def update_employee(emp_id, name, dob_str, age, start_date_str, role, note):
    id_list = main_sheet_obj.col_values(1)
    if emp_id in id_list:
        row_index = id_list.index(emp_id) + 1
        main_sheet_obj.update_cell(row_index, 2, name)
        main_sheet_obj.update_cell(row_index, 3, dob_str)
        main_sheet_obj.update_cell(row_index, 4, age)
        main_sheet_obj.update_cell(row_index, 5, start_date_str)
        main_sheet_obj.update_cell(row_index, 6, role)
        main_sheet_obj.update_cell(row_index, 7, note)
        return True
    return False

def delete_employee(emp_id, name):
    id_list = main_sheet_obj.col_values(1)
    if emp_id in id_list:
        row_index = id_list.index(emp_id) + 1
        main_sheet_obj.delete_rows(row_index)
        try:
            target_sheet = spreadsheet.worksheet(name)
            spreadsheet.del_worksheet(target_sheet)
        except gspread.exceptions.WorksheetNotFound:
            pass
        return True
    return False


# --- 5. GIAO DIỆN ỨNG DỤNG STREAMLIT ---
st.set_page_config(layout="wide")  # Kích hoạt giao diện màn hình rộng để dàn trang biểu đồ
st.title("📊 Hệ thống Quản lý Nhân sự & Hiệu suất Giao dịch")

selected_sheet_name = st.sidebar.selectbox("📂 Chọn trang tính để làm việc:", all_worksheets)
sheet = spreadsheet.worksheet(selected_sheet_name)
st.sidebar.info(f"Đang phân tích dữ liệu tại: **{selected_sheet_name}**")

is_main_sheet = (selected_sheet_name == MAIN_SHEET_NAME)

if is_main_sheet:
    # --- GIAO DIỆN BẢNG TỔNG NHÂN SỰ ---
    tab1, tab2 = st.tabs(["⚙️ Thao tác dữ liệu (CRUD)", "📋 Xem bảng dữ liệu"])

    with tab1:
        action = st.radio("Chọn hành động mong muốn:", ["Thêm nhân viên mới", "Cập nhật thông tin", "Xóa nhân viên"], horizontal=True)
        st.markdown("---")
        
        current_records = main_sheet_obj.get_all_records()
        df_current = pd.DataFrame(current_records)
        
        if action == "Thêm nhân viên mới":
            st.subheader("📥 Đăng ký nhân viên và Khởi tạo không gian riêng")
            with st.form("insert_form", clear_on_submit=True):
                col1, col2 = st.columns(2)
                with col1:
                    name = st.text_input("Họ và tên nhân viên:").strip()
                    dob = st.date_input("Ngày tháng năm sinh:", value=datetime(2000, 1, 1), max_value=datetime.today())
                with col2:
                    start_date = st.date_input("Ngày bắt đầu làm việc:", value=datetime.today())
                    role = st.selectbox("Trạng thái công việc:", ["CTV", "Nhân viên", "Học việc", "Thử việc"])
                note = st.text_area("Ghi chú công việc ban đầu:")
                submit_btn = st.form_submit_button("Lưu & Tự tạo mã ID")
                
            if submit_btn:
                if name == "" or name == MAIN_SHEET_NAME:
                    st.warning("Tên nhân viên không hợp lệ hoặc đang để trống!")
                else:
                    age = datetime.today().year - dob.year
                    dob_str = dob.strftime("%d/%m/%Y")
                    start_date_str = start_date.strftime("%d/%m/%Y")
                    
                    new_id = insert_employee(name, dob_str, age, start_date_str, role, note)
                    st.success(f"🎉 Khởi tạo thành công! Nhân viên **{name}** nhận mã định danh: **{new_id}**")
                    st.rerun()

        elif action == "Cập nhật thông tin":
            st.subheader("📝 Thay đổi thông tin nhân sự")
            if len(current_records) == 0:
                st.info("Chưa có nhân viên nào trong hệ thống để chỉnh sửa.")
            else:
                emp_options = [f"{r['ID']} - {r['Họ và tên']}" for r in current_records]
                selected_emp = st.selectbox("Chọn nhân viên cần sửa thông tin:", emp_options)
                
                selected_id = selected_emp.split(" - ")[0]
                emp_data = df_current[df_current['ID'] == selected_id].iloc[0]
                
                with st.form("update_form"):
                    col1, col2 = st.columns(2)
                    with col1:
                        u_name = st.text_input("Họ và tên:", value=emp_data['Họ và tên']).strip()
                        try:
                            old_dob = datetime.strptime(str(emp_data['Ngày sinh']), "%d/%m/%Y")
                        except:
                            old_dob = datetime(2000, 1, 1)
                        u_dob = st.date_input("Ngày tháng năm sinh:", value=old_dob)
                    with col2:
                        try:
                            old_start = datetime.strptime(str(emp_data['Ngày bắt đầu']), "%d/%m/%Y")
                        except:
                            old_start = datetime.today()
                        u_start_date = st.date_input("Ngày bắt đầu làm việc:", value=old_start)
                        
                        roles = ["CTV", "Nhân viên", "Học việc", "Thử việc"]
                        default_role_idx = roles.index(emp_data['Trạng thái']) if emp_data['Trạng thái'] in roles else 0
                        u_role = st.selectbox("Trạng thái công việc:", roles, index=default_role_idx)
                        
                    u_note = st.text_area("Ghi chú sửa đổi:", value=emp_data['Ghi chú'])
                    update_btn = st.form_submit_button("Cập nhật lại dữ liệu")
                    
                if update_btn:
                    u_age = datetime.today().year - u_dob.year
                    u_dob_str = u_dob.strftime("%d/%m/%Y")
                    u_start_str = u_start_date.strftime("%d/%m/%Y")
                    
                    if update_employee(selected_id, u_name, u_dob_str, u_age, u_start_str, u_role, u_note):
                        st.success(f"💪 Đã cập nhật thành công thông tin của nhân viên mã **{selected_id}**!")
                        st.rerun()
                    else:
                        st.error("Không tìm thấy mã nhân viên trên hệ thống.")

        elif action == "Xóa nhân viên":
            st.subheader("❌ Loại bỏ nhân sự khỏi danh sách")
            if len(current_records) == 0:
                st.info("Hệ thống trống, không có dữ liệu để xóa.")
            else:
                emp_options = [f"{r['ID']} - {r['Họ và tên']}" for r in current_records]
                selected_emp = st.selectbox("Chọn nhân viên muốn xóa khỏi hệ thống:", emp_options)
                
                selected_id = selected_emp.split(" - ")[0]
                selected_name = selected_emp.split(" - ")[1]
                
                st.warning(f"⚠️ Cảnh báo: Hành động này sẽ xóa dòng dữ liệu tổng của **{selected_name}** và XÓA HOÀN TOÀN trang tính riêng mang tên nhân viên này.")
                confirm_check = st.checkbox("Tôi xác nhận muốn xóa vĩnh viễn dữ liệu này.")
                
                if st.button("Tiến hành xóa"):
                    if confirm_check:
                        if delete_employee(selected_id, selected_name):
                            st.success(f"🗑️ Đã xóa hoàn toàn nhân viên **{selected_name}** khỏi hệ thống.")
                            st.rerun()
                        else:
                            st.error("Xảy ra lỗi trong quá trình xóa dữ liệu.")
                    else:
                        st.info("Vui lòng tích vào ô xác nhận trước khi thực hiện xóa.")

    with tab2:
        st.header(f"Bảng hiển thị trực quan: [{selected_sheet_name}]")
        if st.button("🔄 Đồng bộ & Tải lại dữ liệu"):
            st.rerun()
        current_records = main_sheet_obj.get_all_records()
        if len(current_records) > 0:
            st.dataframe(pd.DataFrame(current_records), use_container_width=True)
        else:
            st.info("Trang tính hiện tại chưa có bản ghi nào.")

else:
    # --- GIAO DIỆN HIỆU SUẤT CỦA TỪNG NHÂN VIÊN CỤ THỂ ---
    st.sidebar.success(f"💼 Nhân viên: {selected_sheet_name}")
    
    # Khởi tạo 3 tab chức năng bao gồm tab Phân tích & Đồ thị mới
    tab_trade_input, tab_trade_view, tab_analysis = st.tabs([
        "📥 Nhập số liệu phiên", 
        "📋 Nhật ký giao dịch chi tiết", 
        "📊 Phân tích & Biểu đồ xu hướng"
    ])
    
    with tab_trade_input:
        st.subheader(f"Cập nhật số liệu KPI Phiên giao dịch - [{selected_sheet_name}]")
        
        with st.form("trading_session_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                trade_date = st.date_input("Chọn phiên giao dịch (Ngày):", value=datetime.today())
                trading_value = st.number_input("Giá trị giao dịch (VNĐ):", min_value=0.0, step=10000000.0, format="%.2f")
            with col2:
                new_kh = st.number_input("Số lượng KH mới phát triển:", min_value=0, step=1)
                move_id_kh = st.number_input("Số lượng KH chuyển đổi ID:", min_value=0, step=1)
                
            submit_trade = st.form_submit_button("Ghi nhận & Khóa sổ phiên")
            
        if submit_trade:
            weekday = trade_date.weekday()  # 5: Thứ bảy, 6: Chủ nhật
            
            if weekday == 5 or weekday == 6:
                st.error("❌ Không thể lưu dữ liệu! Ngày bạn chọn là Thứ 7 hoặc Chủ Nhật. Thị trường chứng khoán không giao dịch.")
            elif is_vietnam_holiday(trade_date):
                st.error(f"❌ Không thể lưu dữ liệu! Ngày {trade_date.strftime('%d/%m/%Y')} trùng vào ngày nghỉ lễ của Việt Nam.")
            elif trading_value == 0:
                st.warning("Vui lòng điền Giá trị giao dịch hợp lệ!")
            else:
                with st.spinner("Hệ thống đang tự động tính toán Phí giao dịch (0.25%) và Phí Net (30%)..."):
                    try:
                        # SỬA LỖI: Gán chính xác cặp biến trả về từ hàm tính toán
                        company_fee, net_fee = calculate_trading_fees(trading_value)
                        
                        trade_date_str = trade_date.strftime("%d/%m/%Y")
                        
                        # Khớp đúng cấu trúc mảng để đẩy xuống Google Sheet con
                        trade_row = [trade_date_str, trading_value, company_fee, net_fee, new_kh, move_id_kh]
                        
                        sheet.append_row(trade_row)
                        st.success(f"🎉 Khóa sổ thành công phiên {trade_date_str}! Phí GD Công ty (0.25%): {company_fee:,.0f} VNĐ | Phí Net CTV nhận về (30%): **{net_fee:,.0f} VNĐ**")
                    except Exception as e:
                        st.error(f"Lỗi ghi nhận dữ liệu giao dịch: {e}")

    with tab_trade_view:
        st.subheader(f"Nhật ký hiệu suất: {selected_sheet_name}")
        if st.button("🔄 Làm mới bảng hiệu suất"):
            st.rerun()
            
        trade_records = sheet.get_all_records()
        if len(trade_records) > 0:
            df_trade = pd.DataFrame(trade_records)
            
            # Định dạng hiển thị chuỗi tiền tệ phân tách hàng nghìn cho người dùng dễ nhìn
            df_display = df_trade.copy()
            for col in ['Giá trị giao dịch', 'Phí giao dịch', 'Phí net']:
                if col in df_display.columns:
                    df_display[col] = df_display[col].map(lambda x: f"{x:,.0f}" if isinstance(x, (int, float)) else x)
                    
            st.dataframe(df_display, use_container_width=True)
        else:
            st.info("Nhân viên này hiện tại chưa có dữ liệu giao dịch phát sinh.")

    # 🌟 TAB 3: PHÂN TÍCH CHU KỲ (TUẦN/THÁNG/QUÝ) & 2 BIỂU ĐỒ TĂNG TRƯỞNG
    with tab_analysis:
        st.subheader(f"📈 Trung tâm Phân tích Hiệu suất & Phát triển Khách hàng: {selected_sheet_name}")
        
        trade_records = sheet.get_all_records()
        if len(trade_records) > 0:
            df_raw = pd.DataFrame(trade_records)
            
            # Khởi tạo cột trục thời gian chuẩn hóa để tính toán gom nhóm dữ liệu
            try:
                df_raw['Date_Parsed'] = pd.to_datetime(df_raw['Phiên giao dịch'], format="%d/%m/%Y")
            except Exception:
                st.error("Lỗi: Dữ liệu cột 'Phiên giao dịch' trên Google Sheet không đồng bộ định dạng Ngày/Tháng/Năm!")
                st.stop()
                
            # Ép kiểu dữ liệu số để tránh lỗi tính toán
            df_raw['Phí net'] = pd.to_numeric(df_raw['Phí net'], errors='coerce').fillna(0)
            df_raw['Giá trị giao dịch'] = pd.to_numeric(df_raw['Giá trị giao dịch'], errors='coerce').fillna(0)
            df_raw['KH mới'] = pd.to_numeric(df_raw['KH mới'], errors='coerce').fillna(0)
            df_raw['KH chuyển ID'] = pd.to_numeric(df_raw['KH chuyển ID'], errors='coerce').fillna(0)
            
            # Sắp xếp tăng dần theo dòng thời gian
            df_raw = df_raw.sort_values(by='Date_Parsed')
            
            # --- 3.1 CHỌN CHU KỲ THỜI GIAN CHÈN DỮ LIỆU ---
            time_filter = st.radio("Cấu hình báo cáo tổng hợp theo chu kỳ:", ["Báo cáo Tuần", "Báo cáo Tháng", "Báo cáo Quý"], horizontal=True)
            
            if time_filter == "Báo cáo Tuần":
                df_raw['Thời gian'] = df_raw['Date_Parsed'].dt.to_period('W').astype(str)
            elif time_filter == "Báo cáo Tháng":
                df_raw['Thời gian'] = df_raw['Date_Parsed'].dt.to_period('M').astype(str)
            else:
                df_raw['Thời gian'] = df_raw['Date_Parsed'].dt.to_period('Q').astype(str)
                
            # Gom nhóm tổng lũy kế theo chu kỳ đã chọn
            df_grouped = df_raw.groupby('Thời gian').agg({
                'Phí net': 'sum',
                'Giá trị giao dịch': 'sum',
                'KH mới': 'sum',
                'KH chuyển ID': 'sum'
            }).reset_index()
            
            # --- 3.2 THẺ ĐO LƯỜNG KPI TỔNG HỢP ---
            st.markdown("### 📊 Chỉ số tích lũy tổng quan")
            kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
            with kpi_col1:
                st.metric("Tổng Phí Net nhận về", f"{df_raw['Phí net'].sum():,.0f} VNĐ")
            with kpi_col2:
                st.metric("Tổng Doanh số Giao dịch", f"{df_raw['Giá trị giao dịch'].sum():,.0f} VNĐ")
            with kpi_col3:
                st.metric("Tổng KH Mới phát triển", f"{int(df_raw['KH mới'].sum())} KH")
            with kpi_col4:
                st.metric("Tổng KH Chuyển ID", f"{int(df_raw['KH chuyển ID'].sum())} KH")
                
            st.markdown("---")
            
            # --- 3.3 DÀN TRANG HIỂN THỊ SONG SONG 2 BIỂU ĐỒ ---
            chart_col1, chart_col2 = st.columns(2)
            
            with chart_col1:
                st.markdown(f"#### 📈 1. Xu hướng tăng trưởng doanh thu Phí Net ({time_filter})")
                # Biểu đồ đường (Line Chart) thể hiện mức độ tăng trưởng doanh thu theo thời gian
                df_net_chart = df_grouped.set_index('Thời gian')[['Phí net']]
                st.line_chart(df_net_chart, use_container_width=True)
                
            with chart_col2:
                st.markdown(f"#### 👥 2. Biểu đồ tăng trưởng quy mô số lượng Khách hàng")
                # Tạo cột tổng lượng khách hàng phát triển trong kỳ đó
                df_grouped['Tổng KH mới trong kỳ'] = df_grouped['KH mới'] + df_grouped['KH chuyển ID']
                df_kh_chart = df_grouped.set_index('Thời gian')[['Tổng KH mới trong kỳ']]
                # Biểu đồ cột (Bar Chart) thể hiện sự tăng trưởng tập khách hàng qua từng chu kỳ
                st.bar_chart(df_kh_chart, use_container_width=True)
                
            # --- 3.4 BIỂU ĐỒ TRÒN XEM TỶ LỆ CƠ CẤU KHÁCH HÀNG ---
            st.markdown("### 🍕 3. Phân tích Tỷ lệ cơ cấu Tập khách hàng")
            total_new = df_raw['KH mới'].sum()
            total_moved = df_raw['KH chuyển ID'].sum()
            
            if (total_new + total_moved) > 0:
                # Chuẩn bị bảng dữ liệu cơ cấu phần trăm
                pie_data = pd.DataFrame({
                    'Loại khách hàng': ['Khách hàng mới', 'Khách hàng chuyển ID'],
                    'Số lượng': [total_new, total_moved]
                })
                
                # Vẽ biểu đồ cơ cấu bằng cách tận dụng st.bar_chart dạng hoành độ thành phần (stacked) hoặc hiển thị bảng tỷ lệ trực quan
                c1, c2 = st.columns([1, 2])
                with c1:
                    st.markdown("**Bảng phân rã tỷ lệ chi tiết:**")
                    pct_new = (total_new / (total_new + total_moved)) * 100
                    pct_moved = (total_moved / (total_new + total_moved)) * 100
                    st.write(f"🔹 **Khách hàng mới tinh:** {int(total_new)} người ({pct_new:.1f}%)")
                    st.write(f"🔸 **Khách hàng chuyển ID:** {int(total_moved)} người ({pct_moved:.1f}%)")
                with c2:
                    # Hiển thị biểu đồ thanh ngang so sánh cơ cấu thành phần khách hàng chiếm tỷ lệ cao
                    df_pie_chart = pie_data.set_index('Loại khách hàng')
                    st.bar_chart(df_pie_chart, use_container_width=True)
            else:
                st.info("Chưa có dữ liệu số lượng khách hàng để phân tích cơ cấu tỷ lệ.")
                
            st.markdown("---")
            
            # --- 3.5 BẢNG DỮ LIỆU TỔNG HỢP CUỐI KỲ ---
            st.markdown(f"### 📋 Bảng số liệu tổng hợp lũy kế chi tiết ({time_filter})")
            df_grouped_display = df_grouped.copy()
            df_grouped_display['Phí net'] = df_grouped_display['Phí net'].map('{:,.0f}'.format)
            df_grouped_display['Giá trị giao dịch'] = df_grouped_display['Giá trị giao dịch'].map('{:,.0f}'.format)
            st.dataframe(df_grouped_display, use_container_width=True)
            
        else:
            st.info("Nhân viên này hiện chưa phát sinh bản ghi dữ liệu giao dịch nào để phân tích xu hướng.")