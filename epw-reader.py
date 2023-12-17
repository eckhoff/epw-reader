import pandas as pd
from datetime import datetime


def read_epw_to_dataframe(epw_file_path):
    """
    Reads an EPW file and returns a DataFrame with relevant columns.
    """
    with open(epw_file_path, 'r') as file:
        lines = file.readlines()

    data = [line.strip().split(',') for line in lines[8:]]
    df = pd.DataFrame(data)
    df = df[[0, 1, 2, 3, 6]].copy()  # Selecting year, month, day, hour, and temperature
    df.columns = ['Year', 'Month', 'Day', 'Hour', 'Temperature (°C)']

    df[['Temperature (°C)']] = df[['Temperature (°C)']].apply(pd.to_numeric, errors='coerce')
    df['Temperature (°F)'] = df['Temperature (°C)'] * 9/5 + 32

    
    # Create a list of date strings in 'YYYY-MM-DD' format
    date_strings = df['Year'] + '-' + df['Month'].str.zfill(2) + '-' + df['Day'].str.zfill(2)
    
    # Convert date strings to datetime objects
    date_objects = [datetime.strptime(date_str, '%Y-%m-%d') for date_str in date_strings]
    
    # Calculate the day of the week (0=Monday, 1=Tuesday, ..., 6=Sunday)
    day_of_week = [(date_obj.weekday() + 2) % 7 for date_obj in date_objects]
    
    # Add the 'Day of Week' column to the DataFrame
    df['Day of Week'] = day_of_week

    # Dummy Cell - can use this column to add in RH% or another parameter later
    df['24H/365D'] = 1

    # Change the following columns to int type. reads a string by default.
    df[['Year']] = df[['Year']].apply(pd.to_numeric, errors='coerce')
    df[['Month']] = df[['Month']].apply(pd.to_numeric, errors='coerce')
    df[['Day']] = df[['Day']].apply(pd.to_numeric, errors='coerce')
    df[['Hour']] = df[['Hour']].apply(pd.to_numeric, errors='coerce')

    return df


def write_to_excel(df, excel_file_path):

    with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
        
        # Create Excel Sheet and Tabs
        workbook = writer.book
        worksheet_main = workbook.add_worksheet('Main')
        worksheet_do_not_use = workbook.add_worksheet('DO NOT EDIT -->')
        worksheet_raw_data = workbook.add_worksheet('Raw_Data')
        worksheet_bins = workbook.add_worksheet('Bins')

        # Color Tabs
        worksheet_main.set_tab_color('green')
        worksheet_do_not_use.set_tab_color('red')
        worksheet_raw_data.set_tab_color('red')
        worksheet_bins.set_tab_color('red')

        # Write the raw data in from the EPW to the Raw_Data Tab
        df.to_excel(writer, sheet_name='Raw_Data', index=False)
        
        # Add format condition for titles
        title_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'size': 14,
            'font_color': 'black',
            'border': 2,
            'border_color': 'black',
        })

        # Add format for user input cells
        user_input_cells_format = workbook.add_format({
            'bg_color': '#7eb5ed',
            'align': 'center',
            'valign': 'vcenter',
            'size': 12,
            'font_color': 'black',
            'border': 2,
            'border_color': 'black',
        })

        # Add format for user input cells that are in date (M/D) format
        user_input_date_cells_format = workbook.add_format({
            'bg_color': '#7eb5ed',
            'align': 'center',
            'valign': 'vcenter',
            'size': 12,
            'font_color': 'black',
            'border': 2,
            'border_color': 'black',
            'num_format': 'm/d'
        })

        # Normal Cell Format
        normal_cells_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'size': 12,
            'font_color': 'black',
            'border': 2,
            'border_color': 'black',
        })

        # Add format for data calculation cells
        data_cells_format = workbook.add_format({
            'bg_color': '#f5cc76',
            'align': 'center',
            'valign': 'vcenter',
            'size': 12,
            'font_color': 'black',
            'border': 2,
            'border_color': 'black',
        })

        # Add format for data calculation cells that need to be shown as a percentage
        percent_cells_format = workbook.add_format({
            'bg_color': '#f5cc76',
            'align': 'center',
            'valign': 'vcenter',
            'size': 12,
            'font_color': 'black',
            'border': 2,
            'border_color': 'black',
            'num_format': '0.00%',
        })

        # Add format condition for summary titles
        summary_title_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'size': 12,
            'font_color': 'black',
            'border': 2,
            'border_color': 'black',
        })

        # Add format condition for summary content
        summary_content_format = workbook.add_format({
            'bold': False,
            'bg_color': '#f5cc76',
            'align': 'center',
            'valign': 'vcenter',
            'size': 12,
            'font_color': 'black',
            'border': 2,
            'border_color': 'black',
        })

        # Add bar chart to Main tab
        chart = workbook.add_chart({'type': 'column'})
        chart.add_series({'categories': ['Main', 1, 4, 10, 4],
                          'values': ['Main', 1, 5, 10, 5],
                          'name': 'Temperature Distribution',})
        chart.set_size({'width': 800, 'height': 400})
        chart.set_x_axis({'name': 'Temperature Ranges'})
        chart.set_y_axis({'name': '# of Hours'})
        chart.set_legend({'delete_series': [0]})
        worksheet_main.insert_chart('I5', chart)

        # Set the width of cells in Main tab
        worksheet_main.set_column('A:B', 14)
        worksheet_main.set_column('E:E', 22)
        worksheet_main.set_column('F:F', 17)
        worksheet_main.set_column('G:G', 17)

        # Write the default user input to include or exclude weekends from calculation
        # Include a dropdown list. Only allow "Yes" and "No" in this box
        yes_no_dropdown = ["Yes", "No"]
        worksheet_main.merge_range('A1:B1', 'Include Weekends?', title_format)
        worksheet_main.merge_range('A2:B2', yes_no_dropdown[0], user_input_cells_format)
        worksheet_main.data_validation('A2', {'validate': 'list',
                                              'source': yes_no_dropdown,
                                              'input_message': 'Select Yes or No',
                                              'error_message': 'Please select a valid option'})

        # Write the default user input to specify occupied hours
        # Include a dropdown list. Only allow numbers 0-24
        hours_dropdown = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24]
        worksheet_main.merge_range('A4:B4', 'Occupied Hours', title_format)
        worksheet_main.write('A5', 'Start', normal_cells_format)
        worksheet_main.write('B5', 'End', normal_cells_format)
        worksheet_main.write('A6', hours_dropdown[0], user_input_cells_format)
        worksheet_main.write('B6', hours_dropdown[-1], user_input_cells_format)
        worksheet_main.data_validation('A6:B6', {'validate': 'list',
                                                 'source': hours_dropdown,
                                                 'input_message': 'Select an hour between 0-24',
                                                 'error_message': 'Please select a valid option'})

        # Write the default user input for including breaks
        worksheet_main.merge_range('A8:B8', 'Date Range to Exclude', title_format)
        worksheet_main.write('A9', 'Start', normal_cells_format)
        worksheet_main.write('B9', 'End', normal_cells_format)
        worksheet_main.write('A10', '5/15', user_input_date_cells_format)
        worksheet_main.write('B10', '8/15', user_input_date_cells_format)

        # Add the summary of data
        worksheet_main.merge_range('E14:G14', 'Summary of Conditions Analyzed', summary_title_format)
        worksheet_main.write('E15', 'Days of Week:', summary_title_format)
        worksheet_main.write('E16', 'Hours of Day:', summary_title_format)
        worksheet_main.write('E17', 'Date Range:', summary_title_format)
        worksheet_main.write('E18', 'Total Hours Analyzed:', summary_title_format)
        worksheet_main.merge_range('F15:G15', '=IF(A2="Yes", "Sun-M-Tu-W-Th-F-Sat", "M-Tu-W-Th-F")', summary_content_format)
        worksheet_main.merge_range('F16:G16', '=IF(B6-A6=24, "24 Hours", TEXT(TIME(A6,0,0), "H:MM AM/PM")&" - "&TEXT(TIME(B6,0,0), "H:MM AM/PM"))', summary_content_format)
        worksheet_main.merge_range('F17:G17', '=IF(AND(A10="",B10=""),"Jan 1 - Dec 31",IF(AND(MONTH(A10)=1,DAY(A10)=1),TEXT(B10 + 1, "MMM D")&" - "&"Dec 31","Jan 1 - "&TEXT(A10 - 1, "MMM D")&" and "&TEXT(B10 + 1, "MMM D")&" - "&"Dec 31"))', summary_content_format)
        worksheet_main.merge_range('F18:G18', '=SUM(Raw_Data!L2:L8761)', summary_content_format)
        

        # Write the Main tab spreadsheet output titles
        worksheet_main.write('E1', 'Temp Ranges', title_format)
        worksheet_main.write('F1', '# of Hours', title_format)
        worksheet_main.write('G1', '% of Hours', title_format)
        worksheet_main.write_formula('E2', '="Less than "&I2', data_cells_format)
        worksheet_main.write_formula('E3', '=I2&" to "&J2', data_cells_format)
        worksheet_main.write_formula('E4', '=J2&" to "&K2', data_cells_format)
        worksheet_main.write_formula('E5', '=K2&" to "&L2', data_cells_format)
        worksheet_main.write_formula('E6', '=L2&" to "&M2', data_cells_format)
        worksheet_main.write_formula('E7', '=M2&" to "&N2', data_cells_format)
        worksheet_main.write_formula('E8', '=N2&" to "&O2', data_cells_format)
        worksheet_main.write_formula('E9', '=O2&" to "&P2', data_cells_format)
        worksheet_main.write_formula('E10', '=P2&" to "&Q2', data_cells_format)
        worksheet_main.write_formula('E11', '="Greater than "&Q2', data_cells_format)

        # Write the Main tab total number of hours in each bin
        for i in range(2,12):
            worksheet_main.write_formula(f'F{i}', f'=Bins!C{i}', data_cells_format)

        # Write the Main tab % of hours in each bin to results
        for i in range(2,12):
            worksheet_main.write_formula(f'G{i}', f'=IF(sum(Raw_Data!L2:L8761) <> sum(F2:F11), "ERROR", F{i}/sum(F2:F11))', percent_cells_format)

        # Specify the default temperature bins for cells I2 to Q2 in the Main tab and format
        bin_values = [25, 35, 45, 55, 65, 75, 85, 95, 105]
        worksheet_main.write_row('I2', bin_values, user_input_cells_format)

        # Add a title to the temperature bins in the Main tab and format
        worksheet_main.merge_range('I1:Q1', 'Temperature Bins', title_format)

        # Copy bin values to Bins tab
        bin_columns = ['I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']
        for row_num in range(2, len(bin_values) + 2):
            cell_reference = f'A{row_num}'  # Column A, current row
            new_col = bin_columns[row_num - 2]
            formula = f'=Main!{new_col}2'
            worksheet_bins.write_formula(cell_reference, formula)
        worksheet_bins.write('B1', '24H/365D')
        worksheet_bins.write('C1', 'Output')

        # Sum the number of hours <= 25F for 24H/365D data in the Bins tab
        worksheet_bins.write_formula('B2', '=SUMIFS(Raw_Data!H2:Raw_Data!H8761,Raw_Data!F2:F8761,"<="&A2)')
        # Sum the number of hours 25-35, 35-45, 45-55, 55-65, 65-75, 75-85, 85-95, 95-105 for 24H/365D data in the Bins tab
        for i in range(3,11):
            worksheet_bins.write_formula(f'B{i}', f'=SUMIFS(Raw_Data!H2:H8761,Raw_Data!F2:F8761,">"&A{i-1},Raw_Data!F2:F8761,"<="&A{i})')
        # Sum the number of hours > 105F for 24H/365D data in the Bins tab
        worksheet_bins.write_formula('B11', '=SUMIFS(Raw_Data!H2:H8761,Raw_Data!F2:F8761,">"&A10)')

        # Sum the number of hours <= 25F for output data in the Bins tab
        worksheet_bins.write_formula('C2', '=SUMIFS(Raw_Data!L2:Raw_Data!L8761,Raw_Data!F2:F8761,"<="&A2)')
        # Sum the number of hours 25-35, 35-45, 45-55, 55-65, 65-75, 75-85, 85-95, 95-105 for output data in the Bins tab
        for i in range(3,11):
            worksheet_bins.write_formula(f'C{i}', f'=SUMIFS(Raw_Data!L2:L8761,Raw_Data!F2:F8761,">"&A{i-1},Raw_Data!F2:F8761,"<="&A{i})')
        # Sum the number of hours > 105F for output data in the Bins tab
        worksheet_bins.write_formula('C11', '=SUMIFS(Raw_Data!L2:L8761,Raw_Data!F2:F8761,">"&A10)')
        

        # In Raw Data tab, check if weekends are include or not. Sum total number of hours in year.
        for row_num in range(2, len(df) + 2):  # Start from row 2 (header is row 1)
            cell_reference = f'I{row_num}'  # Column I, current row
            formula = f'=IF(Main!$A$2 = "No", IF(OR(G{row_num} = 3, G{row_num} = 4), 0,1), 1)'
            worksheet_raw_data.write_formula(cell_reference, formula)
        worksheet_raw_data.write('I1', 'Exclude Wkend')

        # In Raw Data tab, sum total hours during occupied hours.
        for row_num in range(2, len(df) + 2):  # Start from row 2 (header is row 1)
            cell_reference = f'J{row_num}'  # Column H, current row
            formula = f'=IF(AND(D{row_num}>=(Main!$A$6 + 1),D{row_num}<=Main!$B$6),1,0)'
            worksheet_raw_data.write_formula(cell_reference, formula)
        worksheet_raw_data.write('J1', 'Occupied Hrs')

        # In Raw Data tab, sum total hours excluding user defined excluded dates
        for row_num in range(2, len(df) + 2):  # Start from row 2 (header is row 1)
            cell_reference = f'K{row_num}'  # Column H, current row
            formula = f'=IF(OR(OR(B{row_num} < MONTH(Main!$A$10), AND(B{row_num} = MONTH(Main!$A$10), C{row_num} < DAY(Main!$A$10))), OR(B{row_num} > MONTH(Main!$B$10), AND(B{row_num} = MONTH(Main!$B$10), C{row_num} > DAY(Main!B$10)))),1,0)'
            worksheet_raw_data.write_formula(cell_reference, formula)
        worksheet_raw_data.write('K1', 'Dates Excluded')

        # In Raw Data tab, sum total hours including all user input variables.
        # including weekends, occupied hours, and excluded date range.
        for row_num in range(2, len(df) + 2):  # Start from row 2 (header is row 1)
            cell_reference = f'L{row_num}'  # Column H, current row
            formula = f'=IF(AND(I{row_num}=1,J{row_num}=1,K{row_num}=1),1,0)'
            worksheet_raw_data.write_formula(cell_reference, formula)
        worksheet_raw_data.write('L1', 'Output')


location = "USA_KS_Hutchinson.Muni.AP.724506_TMY3"
epw_file_path = location + '.epw'
excel_file_path = location + '.xlsx'

df = read_epw_to_dataframe(epw_file_path)
write_to_excel(df, excel_file_path)
