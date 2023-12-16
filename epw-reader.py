import pandas as pd
from datetime import datetime, timedelta


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
    df['24H/365D'] = 1

    df[['Year']] = df[['Year']].apply(pd.to_numeric, errors='coerce')
    df[['Month']] = df[['Month']].apply(pd.to_numeric, errors='coerce')
    df[['Day']] = df[['Day']].apply(pd.to_numeric, errors='coerce')
    df[['Hour']] = df[['Hour']].apply(pd.to_numeric, errors='coerce')

    return df


def write_to_excel(df, excel_file_path):

    with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Raw Data', index=False)

        workbook = writer.book
        worksheet_bins = workbook.add_worksheet('Bins')
        worksheet_sheet1 = workbook.add_worksheet('Sheet1')

        # Write the default user input to include or exclude weekends from calculation
        worksheet_sheet1.write('A1', 'Include Weekends?')
        worksheet_sheet1.write('A2', 'No')

        # Write the default user input to specify occupied hours
        worksheet_sheet1.write('A4', 'Occupied Hours')
        worksheet_sheet1.write('A5', 'Start')
        worksheet_sheet1.write('B5', 'End')
        worksheet_sheet1.write('A6', 1)
        worksheet_sheet1.write('B6', 24)

        # Write the default user input for including breaks
        worksheet_sheet1.write('A8', 'Date Range to Exclude')
        worksheet_sheet1.write('A9', 'Start')
        worksheet_sheet1.write('B9', 'End')
        worksheet_sheet1.write('A10', '5/15')
        worksheet_sheet1.write('B10', '8/15')

        # Write the Main spreadsheet output
        worksheet_sheet1.write('E1', 'Temp Ranges')
        worksheet_sheet1.write('F1', '# of Hours')
        worksheet_sheet1.write('G1', '% of Hours')
        worksheet_sheet1.write_formula('E2', '="Less than "&I2')
        worksheet_sheet1.write_formula('E3', '=I2&" to "&J2')
        worksheet_sheet1.write_formula('E4', '=J2&" to "&K2')
        worksheet_sheet1.write_formula('E5', '=K2&" to "&L2')
        worksheet_sheet1.write_formula('E6', '=L2&" to "&M2')
        worksheet_sheet1.write_formula('E7', '=M2&" to "&N2')
        worksheet_sheet1.write_formula('E8', '=N2&" to "&O2')
        worksheet_sheet1.write_formula('E9', '=O2&" to "&P2')
        worksheet_sheet1.write_formula('E10', '=P2&" to "&Q2')
        worksheet_sheet1.write_formula('E11', '="Greater than "&Q2')
        # Write total number of hours in each bin to results
        for i in range(2,12):
            worksheet_sheet1.write_formula(f'F{i}', f'=Bins!C{i}')
        # Write % of hours in each bin to results
        for i in range(2,12):
            worksheet_sheet1.write_formula(f'G{i}', f'=IF(sum(\'Raw Data\'!L2:L8761) <> sum(F2:F11), "ERROR", F{i}/sum(F2:F11))')

        # Specify the values for cells I2 to Q2 in the "Sheet1"
        bin_values = [25, 35, 45, 55, 65, 75, 85, 95, 105]
        bin_columns = ['I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']
        worksheet_sheet1.write_row('I2', bin_values)
        # Copy bin values to "Bins" sheet
        for row_num in range(2, len(bin_values) + 2):
            cell_reference = f'A{row_num}'  # Column A, current row
            new_col = bin_columns[row_num - 2]
            formula = f'=Sheet1!{new_col}2'
            worksheet_bins.write_formula(cell_reference, formula)
        worksheet_bins.write('B1', '24H/365D')
        worksheet_bins.write('C1', 'Output')

        # Sum the number of hours <= 25F for 24H/365D data
        worksheet_bins.write_formula('B2', '=SUMIFS(\'Raw Data\'!H2:\'Raw Data\'!H8761,\'Raw Data\'!F2:F8761,"<="&A2)')
        # Sum the number of hours 25-35, 35-45, 45-55, 55-65, 65-75, 75-85, 85-95, 95-105 for 24H/365D data
        for i in range(3,11):
            worksheet_bins.write_formula(f'B{i}', f'=SUMIFS(\'Raw Data\'!H2:H8761,\'Raw Data\'!F2:F8761,">"&A{i-1},\'Raw Data\'!F2:F8761,"<="&A{i})')
        # Sum the number of hours > 105F for 24H/365D data
        worksheet_bins.write_formula('B11', '=SUMIFS(\'Raw Data\'!H2:H8761,\'Raw Data\'!F2:F8761,">"&A10)')

        # Sum the number of hours <= 25F for output data
        worksheet_bins.write_formula('C2', '=SUMIFS(\'Raw Data\'!L2:\'Raw Data\'!L8761,\'Raw Data\'!F2:F8761,"<="&A2)')
        # Sum the number of hours 25-35, 35-45, 45-55, 55-65, 65-75, 75-85, 85-95, 95-105 for output data
        for i in range(3,11):
            worksheet_bins.write_formula(f'C{i}', f'=SUMIFS(\'Raw Data\'!L2:L8761,\'Raw Data\'!F2:F8761,">"&A{i-1},\'Raw Data\'!F2:F8761,"<="&A{i})')
        # Sum the number of hours > 105F for output data
        worksheet_bins.write_formula('C11', '=SUMIFS(\'Raw Data\'!L2:L8761,\'Raw Data\'!F2:F8761,">"&A10)')

        # define the "Raw Data" sheet
        worksheet = writer.sheets['Raw Data']

        # Tally weekday only hours
        for row_num in range(2, len(df) + 2):  # Start from row 2 (header is row 1)
            cell_reference = f'I{row_num}'  # Column H, current row
            formula = f'=IF(Sheet1!$A$2 = "No", IF(OR(G{row_num} = 3, G{row_num} = 4), 0,1), 1)'
            worksheet.write_formula(cell_reference, formula)
        worksheet.write('I1', 'Exclude Wkend')

        # Tally occupied hours only
        for row_num in range(2, len(df) + 2):  # Start from row 2 (header is row 1)
            cell_reference = f'J{row_num}'  # Column H, current row
            formula = f'=IF(AND(D{row_num}>=Sheet1!$A$6,D{row_num}<=Sheet1!$B$6),1,0)'
            worksheet.write_formula(cell_reference, formula)
        worksheet.write('J1', 'Occupied Hrs')

        # Tally Dates Excluded hours only
        for row_num in range(2, len(df) + 2):  # Start from row 2 (header is row 1)
            cell_reference = f'K{row_num}'  # Column H, current row
            formula = f'=IF(OR(OR(B{row_num} < MONTH(Sheet1!$A$10), AND(B{row_num} = MONTH(Sheet1!$A$10), C{row_num} < DAY(Sheet1!$A$10))), OR(B{row_num} > MONTH(Sheet1!$B$10), AND(B{row_num} = MONTH(Sheet1!$B$10), C{row_num} > DAY(Sheet1!B$10)))),1,0)'
            worksheet.write_formula(cell_reference, formula)
        worksheet.write('K1', 'Dates Excluded')

        # Tally Output results
        for row_num in range(2, len(df) + 2):  # Start from row 2 (header is row 1)
            cell_reference = f'L{row_num}'  # Column H, current row
            formula = f'=IF(AND(I{row_num}=1,J{row_num}=1,K{row_num}=1),1,0)'
            worksheet.write_formula(cell_reference, formula)
        worksheet.write('L1', 'Output')

        '''
        # Add an Excel formula to the '35 to 45' column
        for row_num in range(2, len(df) + 2):  # Start from row 2 (header is row 1)
            cell_reference = f'I{row_num}'  # Column I, current row
            formula = f'=IF(AND(F{row_num}>Bins!A2, F{row_num}<=Bins!A3), 1, 0)'
            worksheet.write_formula(cell_reference, formula)
        worksheet.write_formula('I8762', 'sum(I2:I8761)')
        worksheet.write('I1', '=Bins!A2&" to "&Bins!A3')
        '''


epw_file_path = 'USA_KS_Hutchinson.Muni.AP.724506_TMY3.epw'
excel_file_path = 'output_file.xlsx'

df = read_epw_to_dataframe(epw_file_path)
write_to_excel(df, excel_file_path)
