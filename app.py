import pandas as pd
import streamlit as st
from datetime import datetime
import io
import warnings

# Suppress warnings
warnings.filterwarnings('ignore')

# Set page config as the first Streamlit command
st.set_page_config(
    page_title="Employee Attendance Tracker", 
    layout="wide",
    page_icon="📊"
)

def check_dependencies():
    """Check if required dependencies are installed"""
    try:
        import openpyxl
        return True
    except ImportError:
        return False

def process_attendance_sheet(sheet_df, sheet_name):
    """Process a single sheet and return structured attendance data"""
    
    # Look for the row that contains dates
    date_row_idx = None
    
    for idx in range(min(5, len(sheet_df))):
        row = sheet_df.iloc[idx]
        date_count = 0
        for cell in row:
            if pd.isna(cell):
                continue
            cell_str = str(cell)
            if (isinstance(cell, datetime) or 
                ('2025-' in cell_str and len(cell_str) > 8) or
                ('2024-' in cell_str and len(cell_str) > 8) or
                ('2023-' in cell_str and len(cell_str) > 8)):
                date_count += 1
                if date_count >= 3:
                    date_row_idx = idx
                    break
        if date_row_idx is not None:
            break
    
    if date_row_idx is None:
        return pd.DataFrame()
    
    # Extract dates from the identified row
    date_columns = []
    actual_dates = []
    
    for col_idx in range(len(sheet_df.columns)):
        if col_idx < 2:  # Skip first 2 columns (Process, EMP ID)
            continue
            
        cell_value = sheet_df.iloc[date_row_idx, col_idx]
        if pd.isna(cell_value) or cell_value == '':
            continue
            
        try:
            if isinstance(cell_value, datetime):
                date_obj = cell_value
            else:
                cell_str = str(cell_value).strip()
                if ' ' in cell_str:
                    date_part = cell_str.split()[0]
                    date_obj = pd.to_datetime(date_part)
                else:
                    date_obj = pd.to_datetime(cell_str)
            
            date_columns.append(col_idx)
            actual_dates.append(date_obj)
            
        except Exception as e:
            continue
    
    if not date_columns:
        return pd.DataFrame()
    
    # Find employee data starting row - handle different sheet structures
    emp_data_start = date_row_idx + 1
    
    # For October sheet, we need to skip the header row that contains "Emp Name"
    if sheet_name == 'October':
        # Check if the next row has "Emp Name" in column C
        if emp_data_start < len(sheet_df):
            next_row_val = sheet_df.iloc[emp_data_start, 2] if len(sheet_df.iloc[emp_data_start]) > 2 else None
            if next_row_val and 'Emp Name' in str(next_row_val):
                emp_data_start += 1
    
    # For November sheet, check if we need to skip header
    elif sheet_name == 'November':
        if emp_data_start < len(sheet_df):
            next_row_val = sheet_df.iloc[emp_data_start, 2] if len(sheet_df.iloc[emp_data_start]) > 2 else None
            if next_row_val and 'Emp Name' in str(next_row_val):
                emp_data_start += 1
    
    # Extract employee attendance data
    attendance_data = []
    
    for idx in range(emp_data_start, len(sheet_df)):
        row = sheet_df.iloc[idx]
        
        # Skip empty rows or rows without proper employee data
        if len(row) < 3 or pd.isna(row.iloc[2]) or str(row.iloc[2]).strip() in ['', 'Emp Name', 'nan', 'None']:
            continue
        
        # Check if this row has any attendance data
        has_attendance_data = False
        for date_col in date_columns:
            if date_col < len(row) and not pd.isna(row.iloc[date_col]) and str(row.iloc[date_col]).strip() not in ['', 'nan']:
                has_attendance_data = True
                break
        
        if not has_attendance_data:
            continue
        
        # Handle different column structures for October vs November
        if sheet_name == 'October':
            # October: Col0=Process, Col1=EMP ID, Col2=Emp Name
            emp_id = str(row.iloc[1]) if len(row) > 1 and not pd.isna(row.iloc[1]) else ''
        else:
            # November: Col0=Process, Col1=EMP ID, Col2=Emp Name  
            emp_id = str(row.iloc[1]) if len(row) > 1 and not pd.isna(row.iloc[1]) else ''
        
        employee_data = {
            'Process': str(row.iloc[0]) if not pd.isna(row.iloc[0]) else 'Unknown',
            'EMP ID': emp_id,
            'Emp Name': str(row.iloc[2]).strip(),
            'Sheet': sheet_name
        }
        
        # Add daily attendance - store dates as strings in consistent format
        for date_col, actual_date in zip(date_columns, actual_dates):
            if date_col < len(row):
                status = row.iloc[date_col]
                if not pd.isna(status) and str(status).strip() not in ['', 'nan']:
                    # Store date as string in consistent format 'YYYY-MM-DD'
                    date_key = actual_date.strftime('%Y-%m-%d')
                    employee_data[date_key] = str(status).strip()
        
        attendance_data.append(employee_data)
    
    return pd.DataFrame(attendance_data)

def calculate_attendance_summary(attendance_df, start_date, end_date):
    """Calculate attendance summary for the selected date range"""
    
    # Define attendance categories
    present_codes = ['W']
    leave_codes = {
        'PL': 'Planned Leave',
        'UPL': 'Unplanned Leave', 
        'SL': 'Sick Leave',
        'CL': 'Casual Leave',
        'Compoff': 'Compensatory Off',
        'Halfday': 'Half Day'
    }
    absent_codes = ['Absconded', 'NCNS']
    resigned_codes = ['Resigned', 'Resgined']
    off_codes = ['OFF']
    
    summary_data = []
    
    for _, employee in attendance_df.iterrows():
        emp_summary = {
            'Process': employee['Process'],
            'EMP ID': employee['EMP ID'],
            'Emp Name': employee['Emp Name'],
            'Sheet': employee['Sheet'],
            'Present': 0,
            'Planned Leave': 0,
            'Unplanned Leave': 0,
            'Sick Leave': 0,
            'Casual Leave': 0,
            'Compensatory Off': 0,
            'Half Day': 0,
            'Absent': 0,
            'Off': 0,
            'Resigned': 0,
            'Total Working Days': 0
        }
        
        # Filter dates in the selected range
        for col in employee.index:
            # Check if column is a date string in format 'YYYY-MM-DD'
            if isinstance(col, str) and len(col) == 10 and col.count('-') == 2:
                try:
                    col_date = pd.to_datetime(col).date()
                    
                    # Check if date is in selected range
                    if start_date <= col_date <= end_date:
                        status = str(employee[col]).strip()
                        
                        if status in present_codes:
                            emp_summary['Present'] += 1
                            emp_summary['Total Working Days'] += 1
                        elif status in leave_codes:
                            leave_type = leave_codes[status]
                            emp_summary[leave_type] += 1
                            if status == 'Halfday':
                                emp_summary['Total Working Days'] += 0.5
                            else:
                                emp_summary['Total Working Days'] += 1
                        elif status in absent_codes:
                            emp_summary['Absent'] += 1
                            emp_summary['Total Working Days'] += 1
                        elif status in resigned_codes:
                            emp_summary['Resigned'] += 1
                        elif status in off_codes:
                            emp_summary['Off'] += 1
                            
                except (ValueError, TypeError) as e:
                    continue
        
        summary_data.append(emp_summary)
    
    return pd.DataFrame(summary_data)

def calculate_combined_summary(summary_df):
    """Calculate combined summary across all months for each employee"""
    
    # Group by employee and sum all attendance counts
    combined_df = summary_df.groupby(['Process', 'EMP ID', 'Emp Name']).agg({
        'Present': 'sum',
        'Planned Leave': 'sum',
        'Unplanned Leave': 'sum',
        'Sick Leave': 'sum',
        'Casual Leave': 'sum',
        'Compensatory Off': 'sum',
        'Half Day': 'sum',
        'Absent': 'sum',
        'Off': 'sum',
        'Resigned': 'sum',
        'Total Working Days': 'sum'
    }).reset_index()
    
    # Add a column to show which sheets the data came from
    sheet_info = summary_df.groupby(['Process', 'EMP ID', 'Emp Name'])['Sheet'].apply(lambda x: ', '.join(sorted(x.unique()))).reset_index()
    combined_df = combined_df.merge(sheet_info, on=['Process', 'EMP ID', 'Emp Name'])
    
    return combined_df

def main():
    st.title("📊 Employee Attendance Tracking System")
    
    # Check dependencies first
    if not check_dependencies():
        st.error("""
        **Missing Required Dependencies**
        
        Please install the required packages:
        ```bash
        pip install openpyxl streamlit pandas
        ```
        """)
        return
    
    st.write("Upload your monthly attendance Excel file to generate reports.")
    
    # File upload
    uploaded_file = st.file_uploader("Choose Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            # Read all sheets
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            
            st.success(f"✅ Found {len(sheet_names)} sheet(s): {', '.join(sheet_names)}")
            
            # Process all sheets
            all_attendance_data = []
            
            for sheet_name in sheet_names:
                with st.spinner(f"📄 Processing {sheet_name}..."):
                    sheet_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
                    processed_data = process_attendance_sheet(sheet_df, sheet_name)
                    if not processed_data.empty:
                        all_attendance_data.append(processed_data)
                        st.success(f"✅ Processed {len(processed_data)} employees from {sheet_name}")
            
            if not all_attendance_data:
                st.error("❌ No valid attendance data found in any sheet.")
                return
            
            # Combine all data
            combined_data = pd.concat(all_attendance_data, ignore_index=True)
            
            st.success(f"✅ Processed attendance data for {len(combined_data)} employees across {len(sheet_names)} sheets")
            
            # Show sample of processed data with date columns
            with st.expander("🔍 View Processed Data Sample"):
                st.write("First 3 employees with first 5 date columns:")
                sample_columns = ['Process', 'EMP ID', 'Emp Name', 'Sheet']
                date_columns = [col for col in combined_data.columns if isinstance(col, str) and len(col) == 10 and col.count('-') == 2]
                if date_columns:
                    display_columns = sample_columns + date_columns[:5]  # Show first 5 date columns
                    st.dataframe(combined_data[display_columns].head(3))
                    
                    st.write(f"Total date columns found: {len(date_columns)}")
                    st.write(f"Date range: {min(date_columns)} to {max(date_columns)}")
                else:
                    st.write("No date columns found in the data")
                    st.write("All columns:", list(combined_data.columns))
            
            # Date range selection
            st.subheader("📅 Select Date Range for Report")
            
            # Get all available dates from the data
            all_dates = []
            date_columns = []
            
            for col in combined_data.columns:
                if isinstance(col, str) and len(col) == 10 and col.count('-') == 2:
                    try:
                        date_obj = pd.to_datetime(col).date()
                        all_dates.append(date_obj)
                        date_columns.append(col)
                    except:
                        continue
            
            if all_dates:
                min_date = min(all_dates)
                max_date = max(all_dates)
                
                st.info(f"📊 Available data from {min_date} to {max_date}")
                
                col1, col2 = st.columns(2)
                with col1:
                    start_date = st.date_input("Start Date", value=min_date, min_value=min_date, max_value=max_date)
                with col2:
                    end_date = st.date_input("End Date", value=max_date, min_value=min_date, max_value=max_date)
                
                if start_date > end_date:
                    st.error("❌ Start date must be before end date")
                else:
                    # Calculate summary
                    with st.spinner("🔄 Generating attendance report..."):
                        summary_df = calculate_attendance_summary(combined_data, start_date, end_date)
                    
                    # Calculate combined summary across all months
                    combined_summary_df = calculate_combined_summary(summary_df)
                    
                    # Display options
                    report_type = st.radio("Select Report Type:", 
                                         ["Combined Report (All Months Total)", "Monthly Breakdown"])
                    
                    if report_type == "Combined Report (All Months Total)":
                        # Display combined summary
                        st.subheader(f"📋 Combined Attendance Summary ({start_date} to {end_date}) - All Months Total")
                        
                        if combined_summary_df.empty:
                            st.warning("No attendance data found for the selected date range.")
                        else:
                            # Filters for combined report
                            col1, col2 = st.columns(2)
                            with col1:
                                process_options = ['All'] + sorted([p for p in combined_summary_df['Process'].unique() if p and p != 'Unknown'])
                                selected_process = st.selectbox("Filter by Process", process_options)
                            with col2:
                                employee_options = ['All'] + sorted([e for e in combined_summary_df['Emp Name'].unique() if e])
                                selected_employee = st.selectbox("Filter by Employee", employee_options)
                            
                            # Apply filters
                            filtered_combined_df = combined_summary_df.copy()
                            if selected_process != 'All':
                                filtered_combined_df = filtered_combined_df[filtered_combined_df['Process'] == selected_process]
                            if selected_employee != 'All':
                                filtered_combined_df = filtered_combined_df[filtered_combined_df['Emp Name'] == selected_employee]
                            
                            # Display combined data
                            st.dataframe(filtered_combined_df, use_container_width=True, height=400)
                            
                            # Download option for combined report
                            csv = filtered_combined_df.to_csv(index=False)
                            st.download_button(
                                label="📥 Download Combined Report as CSV",
                                data=csv,
                                file_name=f"combined_attendance_report_{start_date}_to_{end_date}.csv",
                                mime="text/csv"
                            )
                            
                            # Summary statistics for combined report
                            st.subheader("📈 Combined Summary Statistics")
                            if not filtered_combined_df.empty:
                                col1, col2, col3, col4 = st.columns(4)
                                with col1:
                                    st.metric("Total Employees", len(filtered_combined_df))
                                with col2:
                                    total_present = filtered_combined_df['Present'].sum()
                                    st.metric("Total Present Days", total_present)
                                with col3:
                                    total_pl = filtered_combined_df['Planned Leave'].sum()
                                    st.metric("Total PL Days", total_pl)
                                with col4:
                                    total_upl = filtered_combined_df['Unplanned Leave'].sum()
                                    st.metric("Total UPL Days", total_upl)
                    
                    else:  # Monthly Breakdown
                        # Display monthly breakdown
                        st.subheader(f"📋 Monthly Attendance Breakdown ({start_date} to {end_date})")
                        
                        if summary_df.empty:
                            st.warning("No attendance data found for the selected date range.")
                        else:
                            # Filters for monthly breakdown
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                process_options = ['All'] + sorted([p for p in summary_df['Process'].unique() if p and p != 'Unknown'])
                                selected_process = st.selectbox("Filter by Process", process_options, key="monthly_process")
                            with col2:
                                employee_options = ['All'] + sorted([e for e in summary_df['Emp Name'].unique() if e])
                                selected_employee = st.selectbox("Filter by Employee", employee_options, key="monthly_employee")
                            with col3:
                                sheet_options = ['All'] + sorted(summary_df['Sheet'].unique())
                                selected_sheet = st.selectbox("Filter by Month", sheet_options)
                            
                            # Apply filters
                            filtered_df = summary_df.copy()
                            if selected_process != 'All':
                                filtered_df = filtered_df[filtered_df['Process'] == selected_process]
                            if selected_employee != 'All':
                                filtered_df = filtered_df[filtered_df['Emp Name'] == selected_employee]
                            if selected_sheet != 'All':
                                filtered_df = filtered_df[filtered_df['Sheet'] == selected_sheet]
                            
                            # Display monthly data
                            st.dataframe(filtered_df, use_container_width=True, height=400)
                            
                            # Download option for monthly breakdown
                            csv = filtered_df.to_csv(index=False)
                            st.download_button(
                                label="📥 Download Monthly Report as CSV",
                                data=csv,
                                file_name=f"monthly_attendance_report_{start_date}_to_{end_date}.csv",
                                mime="text/csv"
                            )
            
            else:
                st.error("❌ No valid dates found in the processed data.")
                st.info("Debug info: Showing all column names from processed data:")
                st.write(list(combined_data.columns))
                
        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
            import traceback
            st.code(traceback.format_exc())

if __name__ == "__main__":
    main()
