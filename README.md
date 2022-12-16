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
1. The script clears sheet using the clear_sheet function, clear_sheet requires two inputs.
```Python
clear_sheet(file_path_func,sheet_name_func)
 ```
* file_path_func > The file path 
* sheet_name_func > The sheet name 
2. The script creates a sheet using the create_sheet function,  create_sheet requires three inputs.
```Python
create_sheet(file_path_func,sheet_name_func,input_path_func)
 ```
* file_path_func > The file path 
* sheet_name_func >The sheet name 
* input_path_func >The file path of the input data 

# Technical Details
## Dependencies
```
  - openpyxl=3.0.10
  - pandas=1.4.3
  - pywin32=302
  - xlwings=0.24.9
```
