# exceltools
A small package to manage excell sheets in R

Uses a popup file select window for selecting excel file.
Displays available sheets in the file and lets you select one to load.

Using filepath and sheet arguments allows selecting file and sheet in code

Posibility for easy saving of multiple datatables as sheets to excels file with automatic file ending, so only name is required.

Are not able to append sheets to an existing file, since this will destroy preveous formulas and formatting, so overwrite should be set to FALSE by default.

Any commments are welcome

# Install
devtools::install_github("Hannibal83dk/exceltools")

library(exceltools)

# Example

> library(exceltools)
> 
> table1 <- readsheet()
> 
Choose an excel file to read a sheet

1 Tissue 

2 Blood  

Select: 
> 1
>
> table2 <- readsheet()

> path <- choose.files()
> table3 <- readsheet(path)
> table4 <- readsheet(path, 2)

Choose an excel file to read a sheet

1 Tissue 

2 Blood  

Select: 
> 2
>
> createexcelfile("Tissue_blood_data", table1, table2)
> 
> createexcelfile("Tissue_blood_data", table2, table1, overwrite = TRUE)

