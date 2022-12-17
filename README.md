# Excel Automation
<details><summary>Click for Script</summary>
<p>

```python
from openpyxl import load_workbook
import pandas as pd
import numpy as np
import xlwings as xw
import win32com.client as win32
import pathlib
import sys

def clear_sheet(file_path_func,sheet_name_func): 
    df = pd.read_excel(file_path_func,  sheet_name =sheet_name_func )
    last_col = df.columns.str.contains('Unnamed').tolist().index(True) 
    sheet = wb.sheets[sheet_name_func]
    sheet[1:,0:last_col].clear_contents() 
        
def create_sheet(file_path_func,sheet_name_func,input_path_func):
    
    if isinstance(input_path_func,pd.DataFrame):
        df_input = input_path_func
    else:
        df_input = pd.read_excel(input_path_func)
    
    df_bcf = pd.read_excel(file_path_func,  sheet_name =sheet_name_func )
    mid_point = df_bcf.columns.str.contains('Unnamed').tolist().index(True)
    df_bcf = df_bcf.iloc[:,0:mid_point]
    
    columns = df_bcf.columns 
    df_bcf = df_bcf.iloc[0:0]
    df_bcf.columns =range(df_bcf.shape[1])
    df_input.columns = range(df_input.shape[1]) 
    df_output = pd.concat([df_bcf,df_input])
    df_output.columns = columns 
    
    wb = xw.Book(file_path_func)
    sheet = wb.sheets[sheet_name_func]
    
    formula_dict = {}
    for cell in range(mid_point, mid_point + 50):
        key = sheet[0:1,cell:cell+1].formula
        value = sheet[1:2,cell:cell+1].formula
        formula_dict[key]=value
    
    formula_list = list(formula_dict.keys())
    bcf_list = columns 
    for bcf_header in bcf_list:
        for formula_header in formula_list:
            if bcf_header == formula_header:
                df_output[bcf_header] = formula_dict[bcf_header]
    df_output=df_output.set_index(df_output.columns[0]) 
    sheet.range("A2").options(pd.DataFrame,header = False, expand = 'table',chunksize=1000).value = df_output
```

</p>
</details>

## The Project/ Problem 
In my current role, I was asked to refresh data within excel sheets on a monthly basis. This was done during a time critical period of the month and due to the size of the data and the amount of sheets that required a refresh, the process was eating into my resources.
## The Project Requirements 
- The automation could not replace the current manual process.
- The automation code would have to have little to no maintenance required.
- The automation had to take place outside of excel, e.g., no VBA.
## How it works
### The Sheets
The excel sheets will be in the format below.
![Readme Layout](https://user-images.githubusercontent.com/54468620/207132124-cabf5bcb-ffec-4775-abf6-2601ae18d33c.jpg)

### The Columns
The sheet should contain three distinct column types:
- Data Columns - where data are located.
- Formula Columns - where live formulas used in the sheet are located.
- Formula Bank Columns - where formulas used in the sheet are stored for reference. 

### The Sheet Requirements
The Sheet: 
- Can have any amount of data and formula columns.
- Should have at least one empty column between the last formula column and the first formula bank column.
- The formula column and the corresponding column in the formula bank must have the same name.

![Reedme 5](https://user-images.githubusercontent.com/54468620/207138239-cdc443b9-4445-460d-be38-44a50c4ec18b.jpg)

### The Script 
1. The script clears the sheet using the clear_sheet function, and clear_sheet requires two inputs.
```Python
clear_sheet(file_path_func,sheet_name_func)
 ```
* file_path_func > The file path 
* sheet_name_func > The sheet name 
2. The script creates a sheet using the create_sheet function, and create_sheet requires three inputs:
```Python
create_sheet(file_path_func,sheet_name_func,input_path_func)
 ```
* file_path_func > The file path 
* sheet_name_func > The sheet name 
* input_path_func > The file path of the input data 

# Technical Details
## Dependencies
```
  - openpyxl=3.0.10
  - pandas=1.4.3
  - pywin32=302
  - xlwings=0.24.9
```
## How the Script Works
<details><summary>clear_sheet</summary>
<p>

```python
def clear_sheet(file_path_func,sheet_name_func): 
    df = pd.read_excel(file_path_func,  sheet_name =sheet_name_func )
    last_col = df.columns.str.contains('Unnamed').tolist().index(True) 
    sheet = wb.sheets[sheet_name_func]
    sheet[1:,0:last_col].clear_contents()
```

</p>
</details>
The "clear_sheet" function works by loading the excel sheet into a pandas data frame, then finding the first column header titled "untitled". This header is converted to a coordinate and stored in the variable "last_col".

```python
df = pd.read_excel(file_path_func,  sheet_name =sheet_name_func )
last_col = df.columns.str.contains('Unnamed').tolist().index(True) 
```
Using xlwings, the sheet is cleared up until  "last_col" coordinate, excluding the column headers.
```python
sheet[1:,0:last_col].clear_contents()) 
```
<details><summary>create_sheet</summary>
<p>

```python
def create_sheet(file_path_func,sheet_name_func,input_path_func):
    
    if isinstance(input_path_func,pd.DataFrame):
        df_input = input_path_func
    else:
        df_input = pd.read_excel(input_path_func)
    
    df_bcf = pd.read_excel(file_path_func,  sheet_name =sheet_name_func )
    mid_point = df_bcf.columns.str.contains('Unnamed').tolist().index(True)
    df_bcf = df_bcf.iloc[:,0:mid_point]
    
    columns = df_bcf.columns 
    df_bcf = df_bcf.iloc[0:0]
    df_bcf.columns =range(df_bcf.shape[1])
    df_input.columns = range(df_input.shape[1]) 
    df_output = pd.concat([df_bcf,df_input])
    df_output.columns = columns 
    
    wb = xw.Book(file_path_func)
    sheet = wb.sheets[sheet_name_func]
    
    formula_dict = {}
    for cell in range(mid_point, mid_point + 50):
        key = sheet[0:1,cell:cell+1].formula
        value = sheet[1:2,cell:cell+1].formula
        formula_dict[key]=value
    
    formula_list = list(formula_dict.keys())
    bcf_list = columns 
    for bcf_header in bcf_list:
        for formula_header in formula_list:
            if bcf_header == formula_header:
                df_output[bcf_header] = formula_dict[bcf_header]
    df_output=df_output.set_index(df_output.columns[0]) 
    sheet.range("A2").options(pd.DataFrame,header = False, expand = 'table',chunksize=1000).value = df_output
```
</p>
</details>

"create_sheet" starts by checking if the input variable "input_path_func" is an excel sheet or a pandas dataframe, this allows the function to accept both as input variables and ultimately the input is converted into a dataframe titled "df_input".
 ```python
if isinstance(input_path_func,pd.DataFrame):
        df_input = input_path_func
    else:
        df_input = pd.read_excel(input_path_func) 
```

A dataframe titled "df_bcf" is created from the "file_path_func" and "sheet_name_func" input into the function. The data frame is used to find the first column header titled "untitled", this header is converted to a coordinate and stored in the variable "mid_point". The "mid_point" coordinate is used to cut off "df_bcf" to exclude any data beyond "mid_point".

```python
df_bcf = pd.read_excel(file_path_func,  sheet_name =sheet_name_func )
mid_point = df_bcf.columns.str.contains('Unnamed').tolist().index(True)
df_bcf = df_bcf.iloc[:,0:mid_point] 
```

The script creates a new dataframe titled "df_output" by renaming the columns of "df_bcf" and "df_input" to numbers based on their location, "df_bcf" and "df_input" are then concated together as "df_output"."df_output" columns are renamed to the original "df_bcf" column names. This process essentially creates a dataframe with the data from "df_input" and the column headers from "df_bcf".

```python
columns = df_bcf.columns 
df_bcf = df_bcf.iloc[0:0]
df_bcf.columns =range(df_bcf.shape[1])
df_input.columns = range(df_input.shape[1]) 
df_output = pd.concat([df_bcf,df_input])
df_output.columns = columns
```

A dictionary is created titled "formula_dict", and every formula in the formula bank is added with the key as the formula name and the formula as the value. 

```python
formula_dict = {}
for cell in range(mid_point, mid_point + 50):
    key = sheet[0:1,cell:cell+1].formula
    value = sheet[1:2,cell:cell+1].formula
    formula_dict[key]=value
```

The columns of "df_output" and keys of "formula_dict" are cross-referenced and where the "df_output" column header and the "formula_dict" key matches, the column is filled with the corresponding formula value.

```python
formula_list = list(formula_dict.keys())
bcf_list = columns 
for bcf_header in bcf_list:
     for formula_header in formula_list:
         if bcf_header == formula_header:
             df_output[bcf_header] = formula_dict[bcf_header]
```

"df_output" has the index column removed and is then pasted into the file "file_path_func" on sheet "sheet_name_func".

```python
df_output=df_output.set_index(df_output.columns[0]) 
sheet.range("A2").options(pd.DataFrame,header = False, expand = 'table',chunksize=1000).value = df_output
```

