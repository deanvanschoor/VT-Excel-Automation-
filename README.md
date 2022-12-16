# Excel Automation
## The project/problem 
In my current role I was asked to refresh data within excel sheets on a monthly basis, this was done during a time critical period, due to the size of the data, the amount of sheets that required a refresh the process was eating into my resources.
## The Project Requirements 
- The automation could not replace the current manual process
- The automation code would have to have little to no maintenance required
- The automation had to take place outside of excel. E.G no VBA 
## How it works
### The sheet
The excel sheets will be in the format below.
![Readme Layout](https://user-images.githubusercontent.com/54468620/207132124-cabf5bcb-ffec-4775-abf6-2601ae18d33c.jpg)

### The Columns
The sheet should contain 3 distinct column types
- Data Columns - where data are located
- Formula Columns - where live formulas used in the sheet are located
- Formula Bank Columns - where formulas used in the sheet are stored for reference 

### The Sheet Requirements
The sheet 
- Can have any amount of data and formula columns
- Should have at least one empty column between the last formula column and the first formula bank column
- The formula column and the corresponding column in the formula bank must have the same name

![Reedme 5](https://user-images.githubusercontent.com/54468620/207138239-cdc443b9-4445-460d-be38-44a50c4ec18b.jpg)

### The Script 
1. The script clears the sheet using the clear_sheet function, clear_sheet requires two inputs.
* file_path_func > The file path 
* sheet_name_func > The sheet name 
```Python
def clear_sheet(file_path_func,sheet_name_func): 
    df = pd.read_excel(file_path_func,  sheet_name =sheet_name_func ) 
    last_col = df.columns.str.contains('Unnamed').tolist().index(True) 
    wb = xw.Book(file_path_func)
    sheet = wb.sheets[sheet_name_func]
    sheet[1:,0:last_col].clear_contents()
 ```
2. The script creates a sheet using the create_sheet function,  create_sheet requires three inputs.
* file_path_func > The file path 
* sheet_name_func >The sheet name 
* input_path_func >The file path of the input data 
```Python
def create_sheet(file_path_func,sheet_name_func,input_path_func):
    
    if isinstance(input_path_func,pd.DataFrame):
        df_input = input_path_func
    else:
        df_input = pd.read_excel(input_path_func)
    
    df_bcf = pd.read_excel(file_path_func,  sheet_name =sheet_name_func )
    mid_point = df_bcf.columns.str.contains('Unnamed').tolist().index(True)
    df_bcf = df_bcf.iloc[:,0:mid_point]
    
    columns = df_bcf.columns #### V2 change from
    df_bcf = df_bcf.iloc[0:0]
    df_bcf.columns =range(df_bcf.shape[1])
    df_input.columns = range(df_input.shape[1]) 
    df_output = pd.concat([df_bcf,df_input])
    df_output.columns = columns ###till
    
    wb = xw.Book(file_path_func)
    sheet = wb.sheets[sheet_name_func]
    
    formula_dict = {}
    for cell in range(mid_point, mid_point + 50):
        key = sheet[0:1,cell:cell+1].formula
        value = sheet[1:2,cell:cell+1].formula
        formula_dict[key]=value
    
    formula_list = list(formula_dict.keys())
    bcf_list = columns ###Use columns, V2 change from bcf_list = df_bcf.columns
    for bcf_header in bcf_list:
        for formula_header in formula_list:
            if bcf_header == formula_header:
                df_output[bcf_header] = formula_dict[bcf_header]
    #df_output=df_output.astype('object')  
    df_output=df_output.set_index(df_output.columns[0]) 
    sheet.range("A2").options(pd.DataFrame,header = False, expand = 'table',chunksize=1000).value = df_output
 ```

