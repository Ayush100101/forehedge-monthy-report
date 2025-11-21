import pandas as pd
import streamlit as st
from datetime import datetime
import io

def process_attendance_sheet(sheet_df, sheet_name):
    """Process a single sheet and return structured attendance data"""
    
    # Find the row with dates (look for the first row with date-like values)
    date_row_idx = None
    for idx, row in sheet_df.iterrows():
        # Check if any cell in the row looks like a date
        for cell in row:
            if isinstance(cell, datetime) or (isinstance(cell, str) and any(keyword in str(cell).lower() for keyword in ['2025', '2024', '2023'])):
                date_row_idx = idx
                break
        if date_row_idx is not None:
            break
    
    if date_row_idx is None:
        st.warning(f"Could not find date row in sheet: {sheet_name}")
        return pd.DataFrame()
    
    # Extract dates from the date row
    date_columns = []
    actual_dates = []
    
    for col_idx, col in enumerate(sheet_df.columns):
        cell_value = sheet_df.iloc[date_row_idx, col_idx]
        if isinstance(cell_value, datetime):
            date_columns.append(col_idx)
            actual_dates.append(cell_value)
        elif isinstance(cell_value, str) and any(year in cell_value for year in ['2025', '2024', '2023']):
            try:
                # Try to parse the date string
                date_obj = pd.to_datetime(cell_value)
                date_columns.append(col_idx)
                actual_dates.append(date_obj)
            except:
                continue
    
    if not date_columns:
        st.warning(f"No valid dates found in sheet: {sheet_name}")
        return pd.DataFrame()
    
    # Find the employee data starting row (look for 'Emp Name' column)
    emp_data_start = None
    for idx, row in sheet_df.iterrows():
        if idx <= date_row_idx:
            continue
        if 'Emp Name' in str(row.iloc[2]) if len(row) > 2 else False:
            emp_data_start = idx + 1  # Data starts from next row
            break
    
    if emp_data_start is None:
        # If 'Emp Name' not found, assume data starts after date row
        emp_data_start = date_row_idx + 1
    
    # Extract employee attendance data
    attendance_data = []
    
    for idx in range(emp_data_start, len(sheet_df)):
        row = sheet_df.iloc[idx]
        
        # Skip empty rows or rows without employee name
        if pd.isna(row.iloc[2]) or row.iloc[2] in ['', 'Emp Name']:
            continue
        
        employee_data = {
            'Process': row.iloc[0] if not pd.isna(row.iloc[0]) else '',
            'EMP ID': row.iloc[1] if not pd.isna(row.iloc[1]) else '',
            'Emp Name': row.iloc[2],
            'Sheet': sheet_name
        }
        
        # Add daily attendance
        for date_col, actual_date in zip(date_columns, actual_dates):
            if date_col < len(row):
                status = row.iloc[date_col]
                if not pd.isna(status) and status != '':
                    employee_data[actual_date.date()] = str(status).strip()
        
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
            if isinstance(col, datetime) or (isinstance(col, str) and '-' in str(col)):
                try:
                    if isinstance(col, str):
                        col_date = pd.to_datetime(col).date()
                    else:
                        col_date = col.date() if hasattr(col, 'date') else col
                    
                    # Check if date is in selected range
                    if start_date <= col_date <= end_date:
                        status = str(employee[col]).strip()
                        
                        if status in present_codes:
                            emp_summary['Present'] += 1
                            emp_summary['Total Working Days'] += 1
                        elif status in leave_codes:
                            leave_type = leave_codes[status]
                            emp_summary[leave_type] += 1
                            emp_summary['Total Working Days'] += (0.5 if status == 'Halfday' else 1)
                        elif status in absent_codes:
                            emp_summary['Absent'] += 1
                            emp_summary['Total Working Days'] += 1
                        elif status in resigned_codes:
                            emp_summary['Resigned'] += 1
                        elif status in off_codes:
                            emp_summary['Off'] += 1
                            
                except (ValueError, TypeError):
                    continue
        
        summary_data.append(emp_summary)
    
    return pd.DataFrame(summary_data)

def main():
    st.set_page_config(page_title="Employee Attendance Tracker", layout="wide")
    st.title("📊 Employee Attendance Tracking System")
    
    # File upload
    uploaded_file = st.file_uploader("Upload Attendance Excel File", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            # Read all sheets
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            
            st.success(f"Found {len(sheet_names)} sheet(s): {', '.join(sheet_names)}")
            
            # Process all sheets
            all_attendance_data = []
            
            for sheet_name in sheet_names:
                with st.spinner(f"Processing {sheet_name}..."):
                    sheet_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
                    processed_data = process_attendance_sheet(sheet_df, sheet_name)
                    if not processed_data.empty:
                        all_attendance_data.append(processed_data)
            
            if not all_attendance_data:
                st.error("No valid attendance data found in any sheet.")
                return
            
            # Combine all data
            combined_data = pd.concat(all_attendance_data, ignore_index=True)
            
            st.success(f"Processed attendance data for {len(combined_data)} employees")
            
            # Date range selection
            st.subheader("📅 Select Date Range for Report")
            
            # Get all available dates from the data
            all_dates = []
            for col in combined_data.columns:
                if isinstance(col, datetime) or (isinstance(col, str) and '-' in str(col)):
                    try:
                        if isinstance(col, str):
                            date_obj = pd.to_datetime(col).date()
                        else:
                            date_obj = col.date() if hasattr(col, 'date') else col
                        all_dates.append(date_obj)
                    except:
                        continue
            
            if all_dates:
                min_date = min(all_dates)
                max_date = max(all_dates)
                
                col1, col2 = st.columns(2)
                with col1:
                    start_date = st.date_input("Start Date", value=min_date, min_value=min_date, max_value=max_date)
                with col2:
                    end_date = st.date_input("End Date", value=max_date, min_value=min_date, max_value=max_date)
                
                if start_date > end_date:
                    st.error("Start date must be before end date")
                else:
                    # Calculate summary
                    with st.spinner("Generating attendance report..."):
                        summary_df = calculate_attendance_summary(combined_data, start_date, end_date)
                    
                    # Display summary
                    st.subheader(f"📋 Attendance Summary ({start_date} to {end_date})")
                    
                    # Filters
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        selected_process = st.selectbox("Filter by Process", 
                                                       ['All'] + list(summary_df['Process'].unique()))
                    with col2:
                        selected_employee = st.selectbox("Filter by Employee", 
                                                        ['All'] + list(summary_df['Emp Name'].unique()))
                    
                    # Apply filters
                    filtered_df = summary_df.copy()
                    if selected_process != 'All':
                        filtered_df = filtered_df[filtered_df['Process'] == selected_process]
                    if selected_employee != 'All':
                        filtered_df = filtered_df[filtered_df['Emp Name'] == selected_employee]
                    
                    # Display data
                    st.dataframe(filtered_df, use_container_width=True)
                    
                    # Download option
                    csv = filtered_df.to_csv(index=False)
                    st.download_button(
                        label="Download Report as CSV",
                        data=csv,
                        file_name=f"attendance_report_{start_date}_to_{end_date}.csv",
                        mime="text/csv"
                    )
                    
                    # Summary statistics
                    st.subheader("📈 Summary Statistics")
                    if not filtered_df.empty:
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            st.metric("Total Employees", len(filtered_df))
                        with col2:
                            avg_present = filtered_df['Present'].mean()
                            st.metric("Average Present Days", f"{avg_present:.1f}")
                        with col3:
                            total_pl = filtered_df['Planned Leave'].sum()
                            st.metric("Total PL Days", total_pl)
                        with col4:
                            total_upl = filtered_df['Unplanned Leave'].sum()
                            st.metric("Total UPL Days", total_upl)
            
            else:
                st.error("No valid dates found in the attendance data.")
                
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")

if __name__ == "__main__":
    main()
