import pandas as pd
import xlsxwriter


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
    df['Temp < 23°F'] = (df['Temperature (°F)'] <= 23).astype(int)
    df['23 to 38'] = ((df['Temperature (°F)'] > 23) & (df['Temperature (°F)'] <= 38)).astype(int)
    df['38 to 100'] = 0

    return df


def write_to_excel(df, excel_file_path):

    with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Raw Data', index=False)

        workbook = writer.book
        worksheet_bins = workbook.add_worksheet('Bins')

        # Specify the values for cells A1 to I1 in the "Bins" sheet
        bin_values = [25, 35, 45, 55, 65, 75, 85, 95, 105]
        worksheet_bins.write_column('A1', bin_values)

        worksheet = writer.sheets['Raw Data']
        # Add an Excel formula to the '< 23' column
        for row_num in range(2, len(df) + 2):  # Start from row 2 (header is row 1)
            cell_reference = f'G{row_num}'  # Column G, current row
            formula = f'=IF(F{row_num}<=Bins!A1, 1, 0)'
            worksheet.write_formula(cell_reference, formula)
        # Add an Excel formula to the '23 to 38' column
        for row_num in range(2, len(df) + 2):  # Start from row 2 (header is row 1)
            cell_reference = f'H{row_num}'  # Column H, current row
            formula = f'=IF(AND(F{row_num}>Bins!A1, F{row_num}<=Bins!A2), 1, 0)'
            worksheet.write_formula(cell_reference, formula)


epw_file_path = 'USA_KS_Hutchinson.Muni.AP.724506_TMY3.epw'
excel_file_path = 'output_file.xlsx'
