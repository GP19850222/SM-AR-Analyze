import streamlit as st
import pandas as pd
import altair as alt
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
from st_aggrid.shared import GridUpdateMode
from datetime import datetime
# import numpy as np

# --- Cấu hình trang ---
st.set_page_config(
    page_title="Dashboard Quản Lý Công Nợ ETC",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Hàm hỗ trợ ---
def calculate_age_category_detailed(days_overdue):
    """Phân loại tuổi nợ chi tiết"""
    if days_overdue <= 0:
        return 'Trong hạn'
    elif 1 <= days_overdue <= 30:
        return '1-30'
    elif 31 <= days_overdue <= 60:
        return '31-60'
    elif 61 <= days_overdue <= 90:
        return '61-90'
    else: # days_overdue > 90
        return 'Trên 90'

def format_vnd(amount):
    """Định dạng số tiền sang kiểu VND"""
    if pd.isna(amount) or amount == 0:
        return "0"
    return f"{int(amount):,.0f}" # Không có phần thập phân

def generate_tooltip_html(service_data_df, amount_col_name='amount'):
    """
    Generates an HTML string for tooltips from service breakdown data.
    service_data_df is expected to have 'service_type' and amount_col_name columns.
    """
    if not isinstance(service_data_df, pd.DataFrame) or service_data_df.empty:
        return None
    lines = []
    for _, row in service_data_df.iterrows():
        service = row['service_type']
        amount = row[amount_col_name]
        if pd.notna(amount) and amount > 0: # Only show services with positive amounts
            lines.append(f"{service}: {format_vnd(amount)}")
    return "\n".join(lines) if lines else None

# --- Sidebar ---
st.sidebar.header("📁 Tải Lên Dữ Liệu")
uploaded_file = st.sidebar.file_uploader("Chọn file Excel (.xls, .xlsx)", type=["xls", "xlsx"])

sheet_name_selected = None
xls = None # Initialize xls
if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        if not sheet_names:
            st.sidebar.warning("File Excel không có sheet nào.")
            uploaded_file = None
        else:
            sheet_name_selected = st.sidebar.selectbox("Chọn sheet chứa dữ liệu công nợ:", sheet_names)
    except Exception as e:
        st.sidebar.error(f"Lỗi khi đọc file Excel: {e}")
        uploaded_file = None

# --- Tiêu đề chính ---
st.title("📊 BẢNG DASHBOARD QUẢN LÝ CÔNG NỢ ETC")
st.markdown("---")

if uploaded_file and sheet_name_selected:
    try:
        df_raw = pd.read_excel(uploaded_file, sheet_name=sheet_name_selected)
        st.sidebar.success(f"Đã tải và đọc thành công sheet: '{sheet_name_selected}'")

        # --- Xử lý dữ liệu với Pandas ---
        required_columns = {
            "customer": "KhachHang",      # Tên khách hàng
            "due_date": "NgayDaoHan",     # Ngày đến hạn (phải là định dạng ngày)
            "amount": "SoTienPhaiThu",  # Số tiền còn phải thu
            "service_type": "LoaiHinhDichVu" # Loại hình dịch vụ (R,E,W,P,F)
        }

        missing_cols = []
        actual_columns = {}
        for key, default_name in required_columns.items():
            if default_name in df_raw.columns:
                actual_columns[key] = default_name
            else:
                found_alt = False
                common_alternatives = {
                    "KhachHang": ["khach hang", "ten khach hang", "customer name", "customer","CTY"],
                    "NgayDaoHan": ["ngay dao han", "due date","HẠN TT"],
                    "SoTienPhaiThu": ["so tien phai thu", "amount due", "outstanding amount", "balance","DƯ NỢ"],
                    "LoaiHinhDichVu": ["loai hinh dich vu", "service type", "product type","Loại hình"]
                }
                if default_name in common_alternatives:
                    for alt_name in common_alternatives[default_name]:
                        # Case-insensitive check for alternative column names
                        matching_cols = [col for col in df_raw.columns if col.lower() == alt_name.lower()]
                        if matching_cols:
                            actual_columns[key] = matching_cols[0] # Use the actual casing from file
                            st.sidebar.info(f"Sử dụng cột '{matching_cols[0]}' cho '{default_name}'.")
                            found_alt = True
                            break
                if not found_alt and key in ["customer", "due_date", "amount"]:
                    missing_cols.append(default_name)

        if missing_cols:
            st.error(f"File Excel thiếu các cột bắt buộc sau: {', '.join(missing_cols)}. Vui lòng kiểm tra lại file.")
            st.stop()
        
        df = df_raw.rename(columns={v: k for k, v in actual_columns.items()})

        if 'due_date' not in df.columns:
            st.error("Không tìm thấy cột ngày đáo hạn ('NgayDaoHan' hoặc tương đương).")
            st.stop()
        try:
            df['due_date'] = pd.to_datetime(df['due_date'], errors='coerce')
        except Exception as e:
            st.error(f"Lỗi chuyển đổi cột ngày đáo hạn: {e}. Đảm bảo cột có định dạng ngày tháng hợp lệ.")
            st.stop()

        if 'amount' not in df.columns:
            st.error("Không tìm thấy cột số tiền phải thu ('SoTienPhaiThu' hoặc tương đương).")
            st.stop()
        df['amount'] = pd.to_numeric(df['amount'], errors='coerce').fillna(0)

        if 'service_type' not in df.columns:
            st.warning("Không tìm thấy cột 'LoaiHinhDichVu'. Biểu đồ stacked chart và tooltip chi tiết sẽ không có phân loại dịch vụ.")
            df['service_type'] = 'Không xác định'
        else:
            df['service_type'] = df['service_type'].astype(str).fillna('Không xác định')

        df.dropna(subset=['due_date', 'customer'], inplace=True)
        df = df[df['amount'] > 0].copy()

        if df.empty:
            st.warning("Không có dữ liệu công nợ hợp lệ sau khi xử lý. Vui lòng kiểm tra nội dung file.")
            st.stop()

        current_date = pd.to_datetime(datetime.now().date())
        df['days_overdue'] = (current_date - df['due_date']).dt.days
        df['age_category'] = df['days_overdue'].apply(calculate_age_category_detailed)

        # --- Các chỉ tiêu chính về công nợ (st.metric) ---
        st.subheader("📈 CÁC CHỈ TIÊU CHÍNH VỀ CÔNG NỢ")
        total_receivable = df['amount'].sum()
        overdue_df = df[df['days_overdue'] > 0]
        total_overdue_receivable = overdue_df['amount'].sum()
        overdue_30_days_plus_df = df[df['days_overdue'] > 30]
        total_overdue_30_days_plus = overdue_30_days_plus_df['amount'].sum()

        col1, col2, col3 = st.columns(3)
        col1.metric(label="Tổng Số Dư Công Nợ", value=f"{format_vnd(total_receivable)} VNĐ")
        col2.metric(label="Số Dư Công Nợ Quá Hạn", value=f"{format_vnd(total_overdue_receivable)} VNĐ",
                    delta=f"{format_vnd(total_overdue_receivable - total_receivable)} VNĐ so với tổng nợ", 
                    delta_color="inverse") 
        col3.metric(label="Công Nợ Quá Hạn > 30 Ngày", value=f"{format_vnd(total_overdue_30_days_plus)} VNĐ")
        st.markdown("---")

        # --- Biểu đồ (Altair Chart) ---
        layout_cols = st.columns([6, 4]) 

        with layout_cols[0]:
            st.subheader("📊 Công Nợ Theo Khách Hàng & Loại Hình Dịch Vụ")
            if not df.empty and 'service_type' in df.columns and 'customer' in df.columns:
                chart_stacked_bar = alt.Chart(df).mark_bar().encode(
                    x=alt.X('customer:N', sort='-y', title='Khách Hàng'),
                    y=alt.Y('sum(amount):Q', title='Tổng Dư Nợ (VNĐ)', axis=alt.Axis(format='~s')),
                    color=alt.Color('service_type:N', title='Loại Hình Dịch Vụ'),
                    tooltip=[
                        alt.Tooltip('customer:N', title='Khách Hàng'),
                        alt.Tooltip('service_type:N', title='Loại Dịch Vụ'),
                        alt.Tooltip('sum(amount):Q', title='Dư Nợ', format=',.0f')
                    ]
                ).properties(
                    height=450,
                    title='Tổng công nợ của từng khách hàng theo loại hình dịch vụ'
                )
                st.altair_chart(chart_stacked_bar, use_container_width=True)
            else:
                st.info("Thiếu dữ liệu 'customer' hoặc 'service_type' để tạo biểu đồ stacked bar.")

        with layout_cols[1]:
            st.subheader("🍩 Top 5 Khách Hàng Dư Nợ Lớn Nhất")
            if not df.empty and 'customer' in df.columns:
                customer_total_ar = df.groupby('customer')['amount'].sum().reset_index()
                customer_total_ar_sorted = customer_total_ar.sort_values(by='amount', ascending=False)

                top_5_customers = customer_total_ar_sorted.head(5)
                if len(customer_total_ar_sorted) > 5:
                    other_ar_sum = customer_total_ar_sorted.iloc[5:]['amount'].sum()
                    if other_ar_sum > 0:
                        others_df = pd.DataFrame([{'customer': 'Khách Hàng Khác', 'amount': other_ar_sum}])
                        pie_data = pd.concat([top_5_customers, others_df], ignore_index=True)
                    else:
                        pie_data = top_5_customers
                else:
                    pie_data = top_5_customers

                chart_pie = alt.Chart(pie_data).mark_arc(innerRadius=60, outerRadius=120).encode(
                    theta=alt.Theta(field="amount", type="quantitative", stack=True),
                    color=alt.Color(field="customer", type="nominal", title="Khách Hàng"),
                    tooltip=[
                        alt.Tooltip('customer:N', title='Khách Hàng'),
                        alt.Tooltip('amount:Q', title='Dư Nợ', format=',.0f')
                    ]
                ).properties(
                    height=430,
                    title='Tỷ trọng dư nợ của Top 5 khách hàng'
                )
                st.altair_chart(chart_pie, use_container_width=True)
            else:
                st.info("Thiếu dữ liệu 'customer' để tạo biểu đồ tròn.")

        st.markdown("---")
        number_formatter = JsCode("""
            function formatNumberWithPoint(params) {
                if (params.value == null || isNaN(params.value)) { 
                    return ""; 
                }
                return Number(params.value).toLocaleString('vi-VN', {
                    minimumFractionDigits: 0,
                    maximumFractionDigits: 0
                });
            }""")
        aging_report_containter = st.columns([6, 2]) 

        with aging_report_containter[0]:
            st.subheader("🗓️ Báo Cáo Chi Tiết Công Nợ Phải Thu Theo Tuổi Nợ")

            if not df.empty and 'customer' in df.columns:
                # --- Pivot table for aging report ---
                aging_pivot = pd.pivot_table(
                    df,
                    index='customer',
                    columns='age_category',
                    values='amount',
                    aggfunc='sum',
                    fill_value=0
                )

                age_cols_ordered = ['Trong hạn', '1-30', '31-60', '61-90', 'Trên 90']
                for col in age_cols_ordered:
                    if col not in aging_pivot.columns:
                        aging_pivot[col] = 0
                aging_pivot = aging_pivot[age_cols_ordered] 

                aging_pivot['Dư nợ'] = aging_pivot.sum(axis=1)
                aging_pivot.reset_index(inplace=True)
                aging_pivot = aging_pivot.rename(columns={'customer': 'Khách hàng'})
                
                final_cols_order = ['Khách hàng'] + age_cols_ordered + ['Dư nợ']
                aging_pivot = aging_pivot[final_cols_order]
                
                aging_pivot_sorted = aging_pivot.sort_values(by='Dư nợ', ascending=False).copy()

                # --- NEW: Prepare data for tooltips ---
                # These _tooltip columns are added to the DataFrame to be referenced by tooltipField.
                # They will be explicitly hidden from the grid display later.
                if 'service_type' in df.columns: 
                    service_breakdown_by_age = df.groupby(['customer', 'age_category', 'service_type'])['amount'].sum().reset_index()
                    service_breakdown_total = df.groupby(['customer', 'service_type'])['amount'].sum().reset_index()

                    for age_col in age_cols_ordered:
                        aging_pivot_sorted[f'{age_col}_tooltip'] = None 
                        for i, row in aging_pivot_sorted.iterrows():
                            customer_name = row['Khách hàng']
                            tooltip_df_age = service_breakdown_by_age[
                                (service_breakdown_by_age['customer'] == customer_name) &
                                (service_breakdown_by_age['age_category'] == age_col)
                            ]
                            aging_pivot_sorted.loc[i, f'{age_col}_tooltip'] = generate_tooltip_html(tooltip_df_age, 'amount')

                    aging_pivot_sorted['Dư nợ_tooltip'] = None
                    for i, row in aging_pivot_sorted.iterrows():
                        customer_name = row['Khách hàng']
                        tooltip_df_total = service_breakdown_total[service_breakdown_total['customer'] == customer_name]
                        aging_pivot_sorted.loc[i, 'Dư nợ_tooltip'] = generate_tooltip_html(tooltip_df_total, 'amount')
                # --- END NEW: Prepare data for tooltips ---


                # --- START: MODIFICATION FOR PINNED TOTAL ROW ---
                total_row_data = {'Khách hàng': 'TỔNG CỘNG'}
                numeric_cols_to_sum = age_cols_ordered + ['Dư nợ']
                
                if not aging_pivot_sorted.empty: 
                    total_sum_values = aging_pivot_sorted[numeric_cols_to_sum].sum()
                    for col_name_sum in numeric_cols_to_sum: 
                        total_row_data[col_name_sum] = int(total_sum_values[col_name_sum]) if pd.notna(total_sum_values[col_name_sum]) else 0
                else: 
                    for col_name_sum in numeric_cols_to_sum: 
                        total_row_data[col_name_sum] = 0
                # --- END: MODIFICATION FOR PINNED TOTAL ROW ---

                gb = GridOptionsBuilder.from_dataframe(aging_pivot_sorted) 
                gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=10)
                gb.configure_default_column(filterable=True, sortable=True, resizable=True, aggFunc='sum')
                
                gb.configure_column("Khách hàng", headerName="Khách hàng", width=250, pinned='left',
                                    cellStyle={'textAlign': 'left'})

                currency_cols = age_cols_ordered + ['Dư nợ']
                for col_name in currency_cols: 
                    column_params = {
                        "headerName": col_name.replace("_", " ").title(),
                        "type": ["numericColumn", "numberColumnFilter"],
                        "valueFormatter": number_formatter,
                        "aggFunc": 'sum',
                        "cellStyle": {'textAlign': 'right'}
                    }
                    tooltip_col_name = f'{col_name}_tooltip'
                    if tooltip_col_name in aging_pivot_sorted.columns:
                        column_params["tooltipField"] = tooltip_col_name
                    
                    gb.configure_column(col_name, **column_params)

                # MODIFICATION: Explicitly hide the auxiliary tooltip data columns from display
                # These columns are present in aging_pivot_sorted for tooltipField to access,
                # but they should not be rendered as visible grid columns.
                for col_in_df in aging_pivot_sorted.columns:
                    if col_in_df.endswith('_tooltip'):
                        gb.configure_column(col_in_df, hide=True)


                gridOptions = gb.build()
                
                gridOptions['enableBrowserTooltips'] = True 

                gridOptions['pinnedBottomRowData'] = [total_row_data]
                gridOptions['getRowStyle'] = JsCode("""
                    function(params) {
                        if (params.node.isRowPinned && params.node.rowPinned === 'bottom') {
                            return { 'font-weight': 'bold' };
                        }
                    }
                """)

                AgGrid(
                    aging_pivot_sorted, 
                    gridOptions=gridOptions,
                    height=650,
                    width='100%',
                    fit_columns_on_grid_load=False, 
                    allow_unsafe_jscode=True,
                    enable_enterprise_modules=False, 
                    update_mode=GridUpdateMode.MODEL_CHANGED,
                    key='aging_grid_v_tooltips_hidden' # Changed key
                )
            else:
                st.info("Thiếu dữ liệu 'customer' để tạo báo cáo tuổi nợ.")
            with aging_report_containter[1]:
            # --- Biểu đồ cột cho tổng hợp tuổi nợ (Altair Chart) ---
                aging_summary_data_for_chart = df.groupby('age_category')['amount'].sum().reset_index()
                aging_summary_data_for_chart = aging_summary_data_for_chart.rename(columns={'age_category': 'Tuổi Nợ', 'amount': 'Tổng Số Tiền'})
                
                aging_summary_data_for_chart['Tuổi Nợ'] = pd.Categorical(aging_summary_data_for_chart['Tuổi Nợ'], categories=age_cols_ordered, ordered=True)
                aging_summary_data_for_chart = aging_summary_data_for_chart.sort_values('Tuổi Nợ')


                chart_aging_bar_horizontal = alt.Chart(aging_summary_data_for_chart).mark_bar().encode(
                    y=alt.Y('Tuổi Nợ:N', sort=None, title='Nhóm Tuổi Nợ'),
                    x=alt.X('Tổng Số Tiền:Q', title='Tổng Dư Nợ (VNĐ)', axis=alt.Axis(format='~s')),
                    color=alt.Color('Tuổi Nợ:N', legend=None, scale=alt.Scale(scheme='tableau10')), 
                    tooltip=[
                        alt.Tooltip('Tuổi Nợ:N', title='Nhóm Tuổi Nợ'),
                        alt.Tooltip('Tổng Số Tiền:Q', title='Tổng Dư Nợ', format=',.0f')
                    ]
                ).properties(
                    title='Tổng Quan Công Nợ Phải Thu Theo Tuổi Nợ (Ngang)',
                )
                
                text_horizontal = chart_aging_bar_horizontal.mark_text(
                    align='left',
                    baseline='middle',
                    dx=5 
                ).encode(
                    text=alt.Text('Tổng Số Tiền:Q', format='~s')
                )
                st.altair_chart(chart_aging_bar_horizontal + text_horizontal, use_container_width=True)



    except FileNotFoundError:
        st.error("Lỗi: Không tìm thấy file. Vui lòng đảm bảo đường dẫn file chính xác nếu tải từ server.")
    except pd.errors.EmptyDataError:
        st.error(f"Lỗi: Sheet '{sheet_name_selected}' trống hoặc không có dữ liệu.")
    except KeyError as e:
        st.error(f"Lỗi: Không tìm thấy cột cần thiết trong file Excel: {e}. Vui lòng kiểm tra lại tên cột trong file của bạn so với các tên cột mặc định được mong đợi: {', '.join(required_columns.values())}.")
        if 'df_raw' in locals() and df_raw is not None: 
             st.error(f"Các cột tìm thấy trong sheet '{sheet_name_selected}': {', '.join(df_raw.columns)}")
        else:
             st.error(f"Không thể đọc các cột từ sheet '{sheet_name_selected}'.")
    except ValueError as e:
        st.error(f"Lỗi dữ liệu: {e}. Vui lòng kiểm tra định dạng dữ liệu trong các cột, đặc biệt là cột ngày và số tiền.")
    except Exception as e:
        st.error(f"Đã xảy ra lỗi không mong muốn khi xử lý file: {e}")
        st.exception(e) 
        st.error("Vui lòng kiểm tra lại cấu trúc file Excel, tên các cột và định dạng dữ liệu.")

elif uploaded_file and not sheet_name_selected and xls and xls.sheet_names:
    st.info("Vui lòng chọn một sheet từ file Excel đã tải lên ở thanh bên trái.")
else:
    st.info("👋 Chào mừng! Vui lòng tải lên file Excel báo cáo công nợ để bắt đầu phân tích.")
    st.markdown("""
        #### Hướng dẫn nhanh:
        1.  **Chuẩn bị file Excel:**
            * Đảm bảo file có các cột tối thiểu sau (tên có thể khác, nhưng nội dung tương ứng):
                * `KhachHang`: Tên khách hàng.
                * `NgayDaoHan`: Ngày đến hạn thanh toán (định dạng dd/mm/yyyy, yyyy-mm-dd, v.v.).
                * `SoTienPhaiThu`: Số tiền còn phải thu (dạng số).
                * `LoaiHinhDichVu` (Tùy chọn nhưng cần cho tooltip chi tiết): Loại hình dịch vụ.
            * Dòng tiêu đề (header) nên ở dòng đầu tiên của sheet.
        2.  **Tải file lên:** Sử dụng nút 'Browse files' ở thanh công cụ bên trái (sidebar).
        3.  **Chọn Sheet:** Sau khi tải file thành công, chọn sheet chứa dữ liệu công nợ từ danh sách thả xuống ở sidebar.
        4.  Dashboard sẽ tự động hiển thị các phân tích. Chúc bạn có trải nghiệm tốt!
    """)
