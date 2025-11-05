## 

import pandas as pd
import json
## Read excel from sheet

df = pd.read_excel('test_xl_wrapper.xlsx', sheet_name='Sheet3')

## Transform dataframe to json with columns is first level only
def transform_data_to_json_with_column_is_first_rows(df: pd.DataFrame):
    json_data = []
    columns = df.columns.tolist()
    
    # Create a dictionary for each column using column name as key
    for column in columns:
        column_values = df[column].tolist()
        json_data.append({
            column: column_values
        })
    
    return json_data

from helper.json_helper import convert_numpy_types
ressponse = transform_data_to_json_with_column_is_first_rows(df)
# Convert any numpy types to native Python types
ressponse = convert_numpy_types(ressponse)
# Convert to JSON string
ressponse = json.dumps(ressponse, indent=2, ensure_ascii=False)
print(ressponse)
