# Import Path class to work with files and folders
from pathlib import Path 
# Import pandas module
import pandas as pd

from help_functions import return_format, sum_scraps, get_excel_cell_value

#Scrap defects count - from 'scrap-summ' sheet
COLUMNS_COUNT = 47
NC_RAW_KEY = "NC_raw"
NC_DIVIDED_KEY = "NC_divided"

# Create a Path object pointing to the target folder
folder = Path("./Files")  

sum_scrap_dic = {}
total_produced = 0

# Loop through all items inside the folder
for file in folder.iterdir():
    
    # Create pandas dataframe from current excel's file sheet 'scrap-summ'
    sum_scr_sheet = pd.read_excel(file, sheet_name="scrap-summ")

    # Extract the type of NC column
    nc_type = sum_scr_sheet["Type of NC"]
        
    # Extract the NC column
    nc_column = sum_scr_sheet["NC"]

    # get file produced quantity
    produced_quantity = get_excel_cell_value(exc_file=file, exc_sheet="Таблица за брак", cell_number="G5")

    # get file format
    divisor = return_format(exc_file=file, exc_sheet="Таблица за брак", cell_number="K21")
   
    total_produced += produced_quantity
    
    # iterating over nc_type and nc_column series
    for idx in range(COLUMNS_COUNT):
        # get keys and quantity
        nc_decription = nc_type[idx]
        nc_amount = nc_column[idx]
     
        # add key, value pair if not existing
        if nc_decription not in sum_scrap_dic:
            sum_scrap_dic[nc_decription] = {
                NC_RAW_KEY: 0,
                NC_DIVIDED_KEY: 0
            }
        elif nc_decription == "Technical card- Scrap":
            nc_decription += "-Koch"
            sum_scrap_dic[nc_decription] = {
                NC_RAW_KEY: 0,
                NC_DIVIDED_KEY: 0
            }
            
        # accumulate scrap amount
        sum_scrap_dic[nc_decription][NC_RAW_KEY] += nc_amount

        # filter cells to use divisor on
        if idx < 24 or 35 <= idx <= 39:
            sum_scrap_dic[nc_decription][NC_DIVIDED_KEY] += nc_amount / divisor
        else:
            sum_scrap_dic[nc_decription][NC_DIVIDED_KEY] += nc_amount

#calculate raw and divided scrap
sum_raw_scrap, sum_divided_scrap = sum_scraps(scrap_dic=sum_scrap_dic, key_stop="Damaged or contaminated outer cases", RAW_KEY=NC_RAW_KEY, DIVIDED_KEY=NC_DIVIDED_KEY)

#add new key/value pairs for good pcs and scrap
sum_scrap_dic["Total scrap"] = {
    NC_RAW_KEY: sum_raw_scrap,
    NC_DIVIDED_KEY: round(sum_divided_scrap),
}

sum_scrap_dic["Good Pcs"] = {
    NC_RAW_KEY: total_produced,
    NC_DIVIDED_KEY: total_produced,
}

# Create DataFrame from nested dict
result_df = pd.DataFrame.from_dict(
    sum_scrap_dic,
    orient="index"
)

# Move NC TYPE from index to column
result_df.reset_index(inplace=True)
result_df.rename(columns={"index": "NC TYPE"}, inplace=True)

print(result_df)

# Column names and order
result_df = result_df[["NC TYPE", NC_RAW_KEY, NC_DIVIDED_KEY]]
result_df.rename(columns={
    NC_RAW_KEY: "NC raw",
    NC_DIVIDED_KEY: "NC divided"
}, inplace=True)

#Export to excel
output_file = Path("./nc_summary.xlsx")
result_df.to_excel(output_file, index=False)

print(f"Summary Excel file created at: {output_file}")
