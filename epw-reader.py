import pandas as pd
import matplotlib.pyplot as plt


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

    return df


def create_histogram_image(df, bin_size, image_file_path):
    df['Temperature (°F)'] = df['Temperature (°C)'] * 9/5 + 32
    plt.figure(figsize=(10, 6))
    plt.hist(df['Temperature (°F)'], bins=bin_size, color='blue', edgecolor='black')
    plt.title('Histogram of Temperature')
    plt.xlabel('Temperature (°F)')
    plt.ylabel('Frequency')
    plt.grid(True)
    plt.savefig(image_file_path)
    plt.close()


def write_to_excel_with_histogram(df, excel_file_path, bin_size=10):
    image_file_path = 'histogram.png'
    create_histogram_image(df, bin_size, image_file_path)

    with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Raw Data', index=False)
        # Load workbook and worksheet for adding image
        workbook = writer.book
        worksheet = workbook.add_worksheet('Histogram')
        worksheet.insert_image('B2', image_file_path)


epw_file_path = 'USA_KS_Hutchinson.Muni.AP.724506_TMY3.epw'
excel_file_path = 'output_file.xlsx'
bin_size = 10  # Adjust the bin size here

df = read_epw_to_dataframe(epw_file_path)
write_to_excel_with_histogram(df, excel_file_path, bin_size)
