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

        worksheet_sheet1.write('A2', 'No')
        worksheet_sheet1.write('C3', 1)
        worksheet_sheet1.write('D3', 24)

        # Specify the values for cells A1 to I1 in the "Bins" sheet
        bin_values = [25, 35, 45, 55, 65, 75, 85, 95, 105]
        worksheet_bins.write_column('A1', bin_values)
        # Sum the number of hours <= 25F
        worksheet_bins.write_formula('B1', '=SUMIFS(\'Raw Data\'!G2:\'Raw Data\'!G8761,\'Raw Data\'!F2:F8761,"<="&A1)')
        # Sum the number of hours > 25F & <= 35
        worksheet_bins.write_formula('B2', '=SUMIFS(\'Raw Data\'!G2:G8761,\'Raw Data\'!F2:F8761,">"&A1,\'Raw Data\'!F2:F8761,"<="&A2)')
        # Sum the number of hours > 35F & <= 45
        worksheet_bins.write_formula('B3', '=SUMIFS(\'Raw Data\'!G2:G8761,\'Raw Data\'!F2:F8761,">"&A2,\'Raw Data\'!F2:F8761,"<="&A3)')
        # Sum the number of hours > 45F & <= 55
        worksheet_bins.write_formula('B4', '=SUMIFS(\'Raw Data\'!G2:G8761,\'Raw Data\'!F2:F8761,">"&A3,\'Raw Data\'!F2:F8761,"<="&A4)')
        # Sum the number of hours > 55F & <= 65
        worksheet_bins.write_formula('B5', '=SUMIFS(\'Raw Data\'!G2:G8761,\'Raw Data\'!F2:F8761,">"&A4,\'Raw Data\'!F2:F8761,"<="&A5)')
        # Sum the number of hours > 65F & <= 75
        worksheet_bins.write_formula('B6', '=SUMIFS(\'Raw Data\'!G2:G8761,\'Raw Data\'!F2:F8761,">"&A5,\'Raw Data\'!F2:F8761,"<="&A6)')
        # Sum the number of hours > 75F & <= 85
        worksheet_bins.write_formula('B7', '=SUMIFS(\'Raw Data\'!G2:G8761,\'Raw Data\'!F2:F8761,">"&A6,\'Raw Data\'!F2:F8761,"<="&A7)')
        # Sum the number of hours > 85F & <= 95
        worksheet_bins.write_formula('B8', '=SUMIFS(\'Raw Data\'!G2:G8761,\'Raw Data\'!F2:F8761,">"&A7,\'Raw Data\'!F2:F8761,"<="&A8)')
        # Sum the number of hours > 95F & <= 1055
        worksheet_bins.write_formula('B9', '=SUMIFS(\'Raw Data\'!G2:G8761,\'Raw Data\'!F2:F8761,">"&A8,\'Raw Data\'!F2:F8761,"<="&A9)')
        # Sum the number of hours > 105F
        worksheet_bins.write_formula('B10', '=SUMIFS(\'Raw Data\'!G2:G8761,\'Raw Data\'!F2:F8761,">"&A9)')

        worksheet = writer.sheets['Raw Data']
        # Add an Excel formula to the '< 25' column

        
        # Tally weekday only hours
        for row_num in range(2, len(df) + 2):  # Start from row 2 (header is row 1)
            cell_reference = f'I{row_num}'  # Column H, current row
            formula = f'=IF(Sheet1!$A$2 = "No", IF(OR(G{row_num} = 3, G{row_num} = 4), 0,1), 1)'
            worksheet.write_formula(cell_reference, formula)
        worksheet.write('I1', 'Exclude Wkend')

        # Tally occupied hours only
        for row_num in range(2, len(df) + 2):  # Start from row 2 (header is row 1)
            cell_reference = f'J{row_num}'  # Column H, current row
            formula = f'=IF(AND(D{row_num}>=Sheet1!$C$3,D{row_num}<=Sheet1!$D$3),1,0)'
            worksheet.write_formula(cell_reference, formula)
        worksheet.write('J1', 'Occupied Hrs')

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
