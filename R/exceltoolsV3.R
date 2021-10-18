
#' @import readxl
#' @import openxlsx

library(readxl)
library(openxlsx)

chooseSheet <- function(sheets){
  value = 0

  for (i in sheets) {
    cat(paste(which(sheets == i), i, "\n"))
  }



  while (value == 0) {
    value <- as.numeric(readline(prompt = "Select: "))
  }

  return(value)
}


#' load specific sheets from a chosen excel file
#'
#' Let you chose a file from a popup window, then lists the sheets to choose from
#'
#'
#' @param filepath A path to an excel file
#' @param sheet An integer referring to a specific sheet to open
#' @return A tibble containing the data from the chosen sheet
#' @export
#'
readsheet <- function(filepath = NULL, sheet = NULL){

  if (is.null(filepath)){
    cat("Choose a file to open\n")
    filepath <- file.choose()
  }

  sheets <- excel_sheets(filepath)

  if(length(sheets) < sheet){
    cat("The sheet number provided is larger then the available number of sheets\n")
  }

  if(length(sheets) == 1)
    {
      cat("Only one sheet avalable, selecting sheet 1\n" )
      sheet = 1
    } else if (is.null(sheet) || length(sheets) < sheet)
      {
        cat("Choose a sheet to load \n")
        sheet <- chooseSheet(sheets)
      }

  ##  need only pick first sheet if there only is one, then no selecteion from user

  datasheet <- read_excel(filepath, sheet = sheet)

  return(datasheet)
}




##  appendtoexcellfile <- function(data, file){}

#' Create an excel file and save your tables as sheets in that document
#'
#' As of yet there is no way of giving the sheets other names.
#' The sheets in the resulting document will therefore inherit the variable names given in this function.
#' e.g. table1 will give rise to a sheet in the excel document called table1.
#'
#' @param title A filename for the excel file to be created
#' @param ... A number of dataframes to be used as sheets for the new excel file
#' @param overwrite A logical toggle to choose whether to overwrite a potential existing file
#' @return A tibble containing the data from the chosen sheet
#' @export
#'
createexcelfile <- function(title, ..., overwrite = FALSE){

  wb <- createWorkbook()

  sheets <- list(...)
  sheetnames <- as.list(substitute(list(...)))
  sheetnames <- unlist(sheetnames[2: length(sheetnames)])

  for(i in 1:length(list(...))){
    addWorksheet(wb, sheetnames[[i]])
    writeData(wb, sheetnames[[i]], sheets[[i]])
  }

  saveWorkbook(wb, paste(title, ".xlsx", sep=""), overwrite = overwrite)

  cat(paste("The file ", title, ".xlsx", " has been saved at ", as.character(getwd()), sep = ""))
}



#testsheet1 <- readsheet()
#testsheet2 <- readsheet()
#testsheet3 <- readsheet()

#createexcelfile("excelfiletest3", testsheet1)



#winDialog(type = c("ok", "okcancel", "yesno", "yesnocancel"), message)

#winDialog(type = c("1"), "message")
#winMenuAdd("Testit")
