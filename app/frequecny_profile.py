
# frequency_xcel_data = pd.ExcelFile('../Files/Frequency.xlsx')
frequency_xcel_data = open('../Files/Frequency.xlsx','r')

print(frequency_xcel_data.read())