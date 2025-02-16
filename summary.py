import pandas as pd
import numpy as np

def clean_currency(value):
    """Clean currency strings to numeric values"""
    if isinstance(value, str):
        return float(value.replace('$', '').replace(',', '').strip())
    return float(value)

def process_timesheet(file_path):
    try:
        # Read Excel file
        df = pd.read_excel(file_path)
        
        # Process data into required format
        processed_data = []
        
        # Group by INV # and job
        current_inv = '2277'  # You might want to make this dynamic
        
        # Group by employee and activity
        groups = df.groupby(['Last Name', 'First Name', 'Dist Job Desc', 'Dist Activity Code', 'Dist Activity Desc'])
        
        for name, group in groups:
            last_name, first_name, job_desc, activity_code, activity_desc = name
            employee_name = f"{first_name} {last_name}"
            job_number = activity_code.split(';')[0] if ';' in activity_code else activity_code
            week_ending = group['Date'].max()
            
            # Calculate regular hours and rate
            regular_data = group[group['Earning Desc'] == 'Regular']
            regular_hours = regular_data['Hours'].sum()
            regular_rate = clean_currency(regular_data['Rate'].iloc[0]) if not regular_data.empty else 0
            
            # Calculate overtime hours and rate
            overtime_data = group[group['Earning Desc'] == 'Overtime']
            overtime_hours = overtime_data['Hours'].sum()
            overtime_rate = clean_currency(overtime_data['Rate'].iloc[0]) if not overtime_data.empty else regular_rate * 1.5
            
            # Create base data
            base_data = {
                'INV #': current_inv,
                'EMPLOYEE': employee_name,
                'JOB NAME': job_desc,
                'Activity Code': activity_code,
                'Activity Description': activity_desc,
                'JOB NUMBER': job_number,
                'WEEK ENDING': week_ending
            }
            
            # Create overtime entry first
            overtime_entry = base_data.copy()
            overtime_entry.update({
                'PAY TYPE': 'Overtime',
                'HOURS': f"{overtime_hours:.2f}",
                'BURDENED RATE': overtime_rate,
                'TOTAL': overtime_hours * overtime_rate
            })
            processed_data.append(overtime_entry)
            
            # Create regular entry second
            regular_entry = base_data.copy()
            regular_entry.update({
                'PAY TYPE': 'Regular',
                'HOURS': f"{regular_hours:.2f}",
                'BURDENED RATE': regular_rate,
                'TOTAL': regular_hours * regular_rate
            })
            processed_data.append(regular_entry)
        
        # Create DataFrame
        result_df = pd.DataFrame(processed_data)
        
        # Format currency columns
        result_df['TOTAL'] = result_df['TOTAL'].apply(lambda x: f" ${x:,.2f}")
        result_df['BURDENED RATE'] = result_df['BURDENED RATE'].apply(lambda x: f"${x:,.2f}")
        
        # Sort the DataFrame exactly as in the example
        result_df = result_df.sort_values(['INV #', 'EMPLOYEE', 'Activity Code', 'PAY TYPE'])
        
        # Ensure overtime rows come before regular rows within each group
        result_df['sort_order'] = (result_df['PAY TYPE'] == 'Regular').astype(int)
        result_df = result_df.sort_values(['INV #', 'EMPLOYEE', 'Activity Code', 'sort_order'])
        result_df = result_df.drop('sort_order', axis=1)
        
        return result_df
        
    except Exception as e:
        print(f"Error processing file: {e}")
        import traceback
        print(f"Full error details:\n{traceback.format_exc()}")
        return None

def export_timesheet(df, output_path='timesheet_summary.xlsx'):
    try:
        if df is None:
            return
            
        # Export to Excel
        df.to_excel(output_path, index=False)
        print(f"Timesheet summary exported to {output_path}")
            
    except Exception as e:
        print(f"Error exporting file: {e}")

def main():
    file_path = r'C:\Users\ryanblock\Desktop\summaryapp\excelfilehere\7 twenty 4 Timecard W-E 2.9.25.xlsx'
    print(f"Processing file: {file_path}")
    formatted_df = process_timesheet(file_path)
    if formatted_df is not None:
        export_timesheet(formatted_df)

if __name__ == "__main__":
    main()