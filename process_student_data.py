import pandas as pd

def process_student_attendance(filepath):
    """
    Process the Excel file to extract student attendance data.
    Returns a list of rows containing [subject_name, average_students].
    """
    try:
        print("\n=== Processing Student Attendance Data ===")
        df = pd.read_excel(
            filepath,
            sheet_name="Analiza nastave",
            header=None,
            engine='openpyxl'
        )
        
        # Initialize result data
        table_data = []
        
        # Find the header row that contains "Prosečan broj studenata"
        header_row = None
        student_col = None
        
        for idx, row in df.iterrows():
            row_values = [str(val).strip() if pd.notna(val) else "" for val in row]
            if 'Prosečan broj studenata' in row_values:
                header_row = idx
                student_col = row_values.index('Prosečan broj studenata')
                break
        
        if header_row is None:
            raise ValueError("Could not find header row with 'Prosečan broj studenata'")
            
        # Initialize headers
        table_data.append(["Predmeti", "Prosečan broj studenata"])
        
        # Initialize tracking variables
        current_location = None
        current_year = None
        current_month = None
        months = ['jan', 'feb', 'mar', 'apr', 'maj', 'jun', 'jul', 'avg', 'sep', 'okt', 'nov', 'dec']
        
        # Process data rows
        for idx, row in df.iterrows():
            if idx <= header_row:  # Skip header rows
                continue
                
            row_values = [str(val).strip() if pd.notna(val) else "" for val in row]
            if not any(row_values):  # Skip empty rows
                continue
            
            first_value = next((val for val in row_values if val), "")
            
            # Check for location (e.g., 'Niš')
            if first_value == 'Niš':
                current_location = first_value
                continue
                
            # Check for year
            if first_value.isdigit() and len(first_value) == 4:
                current_year = first_value
                continue
                
            # Check for month
            if first_value.lower() in months:
                current_month = first_value
                continue
            
            # Process subject rows
            if first_value and not first_value.isdigit() and first_value.lower() not in months:
                subject_name = first_value
                
                # Look for student count in this row at the student_col position
                if student_col is not None and student_col < len(row_values):
                    try:
                        val = row_values[student_col]
                        num = float(str(val).replace(',', '.'))
                        if 0 < num < 1000:  # Reasonable range for student count
                            prefix = f"{current_location}-" if current_location else ""
                            full_name = f"{prefix}{current_year}-{current_month}-{subject_name}"
                            table_data.append([full_name, str(int(num) if num.is_integer() else num)])
                    except (ValueError, AttributeError):
                        pass
        
        if len(table_data) <= 1:  # Only headers
            raise ValueError("No student attendance data found")
            
        print("\nProcessed Student Attendance Data:")
        for row in table_data:
            print(row)
            
        return table_data

    except Exception as e:
        print(f"Error processing student attendance data: {e}")
        import traceback
        traceback.print_exc()
        return None 