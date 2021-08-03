
#' @import readxl
#' @import openxlsx

chooseSheet <- function(sheets){
  value = 0
  cat("Choose a sheet to add concentrations \n")
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
#' @return A tibble containing the data from the chosen sheet
#' @export
#'
readsheet <- function(){
  cat("Choose a file to add concentrations")
  file <- file.choose()

  sheets <- excel_sheets(file)

  sheet <- chooseSheet(sheets)

  ##  need only pick first sheet if there only is one, then no selecteion from user

  datasheet <- read_excel(file, sheet = sheet)

  return(datasheet)
}




##  appendtoexcellfile <- function(data, file){}

#' load specific sheets from a chosen excell file
#'
#' Let you chose a file from a popup window, then lists the sheets to choose from
#'
#' @param title A filename for the excel file to be created
#' @param ... A number of dataframes to be used as sheets for the new excel file
#' @param overwrite A logical toggle to choose whether to overwrite a potential existing file
#' @return A tibble containing the data from the chosen sheet
#' @export
#'
createexcelfile <- function(title, ..., overwrite = TRUE){

  wb <- createWorkbook()

  sheets <- list(...)
  sheetnames <- as.list(substitute(list(...)))
  sheetnames <- unlist(sheetnames[2: length(sheetnames)])

  for(i in 1:3){
    addWorksheet(wb, sheetnames[[i]])
    writeData(wb, sheetnames[[i]], sheets[[i]])
  }

  saveWorkbook(wb, paste(title, ".xlsx", sep=""), overwrite = overwrite)

  cat(paste("The file ", title, ".xlsx", " has been saved at ", as.character(getwd()), sep = ""))
}

# testsheet1 <- readsheet()
# testsheet2 <- readsheet()
# testsheet3 <- readsheet()

# createexcellfile("excelfiletest2", testsheet1, testsheet2, testsheet3, overwrite = FALSE)



#winDialog(type = c("ok", "okcancel", "yesno", "yesnocancel"), message)

#winDialog(type = c("1"), "message")
#winMenuAdd("Testit")
