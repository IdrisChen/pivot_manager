import pandas as pd

def transform_and_save_excel(input_file_path,output_file_path):
    # Load the Excel file
    xls = pd.ExcelFile(input_file_path)

    # Load the 'Raw Data Report' sheet
    raw_data = pd.read_excel(xls, 'Raw Data Report')

    filtered_data_tw = raw_data[raw_data['行銷活動名稱'].str.contains('台灣', na=False)]
    
    aggregated_data = filtered_data_tw.groupby('廣告名稱').agg({
        '曝光次數': 'sum',
        '連結點擊次數': 'sum',
        '購買次數': 'sum',
        '購買轉換值': 'sum',
        '加到購物車次數': 'sum',
        '花費金額 (TWD)': 'sum'
    }).reset_index()

    # Handle NaNs in '購買次數' and '購買轉換值'
    aggregated_data[['購買次數', '購買轉換值']] = aggregated_data[['購買次數', '購買轉換值']].fillna(0)

    # Saving the transformed data back into the Excel file
    with pd.ExcelWriter(output_file_path) as writer:
        aggregated_data.to_excel(writer, sheet_name='工作表1', index=False)

    aggregated_data = filtered_data_tw.groupby('天數').agg({
        '曝光次數': 'sum',
        '連結點擊次數': 'sum',
        '購買次數': 'sum',
        '購買轉換值': 'sum',
        '加到購物車次數': 'sum',
        '花費金額 (TWD)': 'sum'
    }).reset_index()

    # Handle NaNs in '購買次數' and '購買轉換值'
    aggregated_data[['購買次數', '購買轉換值']] = aggregated_data[['購買次數', '購買轉換值']].fillna(0)

    # Saving the transformed data back into the Excel file
    with pd.ExcelWriter(output_file_path,  mode='a', if_sheet_exists='new') as writer:
        aggregated_data.to_excel(writer, sheet_name='工作表2', index=False)

############
        
    filtered_data_tw = raw_data[raw_data['行銷活動名稱'].str.contains('港澳', na=False)]
    # Aggregate the data similar to the format of '工作表1'
    aggregated_data = filtered_data_tw.groupby('廣告名稱').agg({
        '曝光次數': 'sum',
        '連結點擊次數': 'sum',
        '購買次數': 'sum',
        '購買轉換值': 'sum',
        '加到購物車次數': 'sum',
        '花費金額 (TWD)': 'sum'
    }).reset_index()

    # Handle NaNs in '購買次數' and '購買轉換值'
    aggregated_data[['購買次數', '購買轉換值']] = aggregated_data[['購買次數', '購買轉換值']].fillna(0)

    # Saving the transformed data back into the Excel file
    with pd.ExcelWriter(output_file_path,  mode='a', if_sheet_exists='new') as writer:
        aggregated_data.to_excel(writer, sheet_name='工作表3', index=False)

    aggregated_data = filtered_data_tw.groupby('天數').agg({
        '曝光次數': 'sum',
        '連結點擊次數': 'sum',
        '購買次數': 'sum',
        '購買轉換值': 'sum',
        '加到購物車次數': 'sum',
        '花費金額 (TWD)': 'sum'
    }).reset_index()

    # Handle NaNs in '購買次數' and '購買轉換值'
    aggregated_data[['購買次數', '購買轉換值']] = aggregated_data[['購買次數', '購買轉換值']].fillna(0)

    # Saving the transformed data back into the Excel file
    with pd.ExcelWriter(output_file_path,  mode='a', if_sheet_exists='new') as writer:
        aggregated_data.to_excel(writer, sheet_name='工作表4', index=False)

# Example usage
input_file_path = 'input_ads_data.xlsx'  # Replace with your file path
output_file_path = 'output_ads_data.xlsx' # Replace with your desired output path
transform_and_save_excel(input_file_path, output_file_path)
