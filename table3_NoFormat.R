############################################################
## Codename - table3_wrangling                            ##
## Data Release - 
## Original Authors - Marios Alexandropoulos              ##
## Orginal Date - December 2018                           ##
## Latest Update Author -   Marios Alexandropoulos        ##
## Latest Update Date -   18/12/18                        ##
## Updates to script (if any):
##                                                        ##
## Type - Extract/model                                   ##
## Written/run on - R Studio Server                       ##
## Type - Preparation                                     ##
## Written/run on - R Studio Server                       ##
## R version - 3.2.3 (2015-12-10)                         ##
## Versions of packages -                                 ##
## tidyverse_1.1.1                                        ##
##                                                        ##
## Description -  Template for producing Table 3: 
## Impact of the data refresh on the previously published
## provisional Standardised Mortality Ratios for all 
## hospitals in Scotland   
##                                                        ##
##                                                        ## 
##                                                        ##
## Approximate run time: 1min                             ## 
############################################################

################################## 
##################################
### SECTION 1 - HOUSE KEEPING ----
##################################
##################################

install.packages("openxlsx") #Read, Write and Edit XLSX Files
library("openxlsx")


# Create a workbook
wb <- loadWorkbook("Table3_no_formatting.xlsx")


# Add and format title

style_title <- createStyle(fontSize = 12, textDecoration = "bold", wrapText = FALSE, fontName = "Arial")

writeData(wb, "HSMR", "Table 3: Impact of the data refresh on the previously published provisional",
          startCol = 1, startRow = 2, rowNames = FALSE)

writeData(wb, "HSMR", "Standardised Mortality Ratios for all hospitals in Scotland; Oct-Dec 2017",
          startCol = 1, startRow = 3, rowNames = FALSE)

addStyle(wb, "HSMR", style_title, rows = 2, cols = 1)
addStyle(wb, "HSMR", style_title, rows = 3, cols = 1)


# Create and format table's column names
writeData(wb, "HSMR", "Oct-Dec 2017",
          startCol = 2, startRow = 5, rowNames = FALSE)

writeData(wb, "HSMR", "HSMR",
          startCol = 2, startRow = 6, rowNames = FALSE)

writeData(wb, "HSMR", "as at 15/05/2018",
          startCol = 2, startRow = 7, rowNames = FALSE)

style <- createStyle(fontSize = 12, textDecoration = "bold", wrapText = FALSE, fontName = "Arial")
addStyle(wb, "HSMR", style, rows = 4:7, cols = 2)


writeData(wb, "HSMR", "Oct-Dec 2017",
          startCol = 3, startRow = 5, rowNames = FALSE)

writeData(wb, "HSMR", "Patients",
          startCol = 3, startRow = 6, rowNames = FALSE)

writeData(wb, "HSMR", "as at 15/05/2018",
          startCol = 3, startRow = 7, rowNames = FALSE)

style <- createStyle(fontSize = 12, textDecoration = "bold", wrapText = FALSE, fontName = "Arial")
addStyle(wb, "HSMR", style, rows = 4:7, cols = 3)



writeData(wb, "HSMR", "Oct-Dec 2017",
          startCol = 4, startRow = 5, rowNames = FALSE)

writeData(wb, "HSMR", "HSMR",
          startCol = 4, startRow = 6, rowNames = FALSE)

writeData(wb, "HSMR", "as at 14/08/2018",
          startCol = 4, startRow = 7, rowNames = FALSE)

style <- createStyle(fontSize = 12, textDecoration = "bold", wrapText = FALSE, fontName = "Arial")
addStyle(wb, "HSMR", style, rows = 4:7, cols = 4)


writeData(wb, "HSMR", "Oct-Dec 2017",
          startCol = 5, startRow = 5, rowNames = FALSE)

writeData(wb, "HSMR", "Patients",
          startCol = 5, startRow = 6, rowNames = FALSE)

writeData(wb, "HSMR", "as at 14/08/2018",
          startCol = 5, startRow = 7, rowNames = FALSE)

style <- createStyle(fontSize = 12, textDecoration = "bold", wrapText = FALSE, fontName = "Arial")
addStyle(wb, "HSMR", style, rows = 4:7, cols = 5)

setColWidths(wb, "HSMR", cols = 2:5, widths = "auto")

# Add and format source title
source_title <- createStyle(fontSize = 8, textDecoration = "bold", wrapText = FALSE, fontName = "Arial")

writeData(wb, "HSMR", "Source: ISD Scotland (SMR01) linked dataset.Reflects the completeness of SMR01 submissions to ISD for individual hospitals as of 12th July 2018.",
          startCol = 1, startRow = 54, rowNames = FALSE)


addStyle(wb, "HSMR", source_title, rows = 54, cols = 1)

# Format table's borders
borderstyle <- createStyle(fontSize = 12, fontName = "Arial",
                           border= "TopBottomLeftRight")

addStyle(wb, "HSMR", borderstyle, rows = 8:52, cols = 1, gridExpand = FALSE)

borderstyle1 <- createStyle(fontSize = 12, fontName = "Arial",
                            border= "TopBottomLeftRight", numFmt = "0.00" )

addStyle(wb, "HSMR", borderstyle1, rows = 8:52, cols = c(2,4), gridExpand = FALSE )


borderstyle2 <- createStyle(fontSize = 12, fontName = "Arial",
                            border= "TopBottomLeftRight", numFmt = "COMMA" )

addStyle(wb, "HSMR", borderstyle2, rows = 8:51, cols = c(3,5), gridExpand = TRUE )

# Format HB and hospital column AND Scotland row

boardstyle <- createStyle(fontSize = 12, textDecoration = "bold", wrapText = FALSE, halign = "left", fontName = "Arial", border= "TopBottomLeftRight")
addStyle(wb, "HSMR", boardstyle, rows = c(8,11,13,15,17,19,22,27,32,36,40,42,44,46,50,52), cols = 1, gridExpand = TRUE)

Scotstyle <- createStyle(fontSize = 12, textDecoration = "bold", wrapText = FALSE, halign = "right", fontName = "Arial",numFmt = "0.00", border= "TopBottomLeftRight")
addStyle(wb, "HSMR", Scotstyle, rows = 52, cols = c(2,4), gridExpand = TRUE)

Scotstyle1 <- createStyle(fontSize = 12, textDecoration = "bold", wrapText = FALSE, halign = "right", fontName = "Arial",numFmt = "COMMA", border= "TopBottomLeftRight")
addStyle(wb, "HSMR", Scotstyle1, rows = 52, cols = c(3,5), gridExpand = TRUE)

# Format fontcolour
colourstyle <- createStyle(fontSize = 12, fontName = "Arial",
                           textDecoration = "bold", halign = "left", bgFill = "#FFC7CE", fontColour = "white", border= "TopBottomLeftRight")

addStyle(wb, "HSMR", colourstyle, rows = 5:7, cols = 1:5, gridExpand = TRUE)

colourstyle1 <- createStyle(fontSize = 10, fontName = "Arial",
                            textDecoration = "bold", halign = "left", bgFill = "#FFC7CE", fontColour = "white", border= "TopBottomLeftRight")

addStyle(wb, "HSMR", colourstyle1, rows = 7, cols = 1:5, gridExpand = TRUE)



style <- createStyle(fontSize = 12, fontName = "Arial",
                     textDecoration = "bold", halign = "left", bgFill = "#FFC7CE", fontColour = "white", border= "TopBottomLeftRight")

addStyle(wb, "HSMR", colourstyle, rows = 5:7, cols = 1:5, gridExpand = TRUE)

colourstyle1 <- createStyle(fontSize = 10, fontName = "Arial",
                            textDecoration = "bold", halign = "left", bgFill = "#FFC7CE", fontColour = "white", border= "TopBottomLeftRight")

addStyle(wb, "HSMR", colourstyle1, rows = 7, cols = 1:5, gridExpand = TRUE)




borderstyle3 <- createStyle(fontSize = 12, fontName = "Arial",
                            border= "TopBottomLeftRight", numFmt = "0.00" )

addStyle(wb, "HSMR", borderstyle3, rows = 8:51, cols = 2, gridExpand = FALSE )

borderstyle4 <- createStyle(fontSize = 12, fontName = "Arial",
                            border= "TopBottomLeftRight", numFmt = "0.00" )

addStyle(wb, "HSMR", borderstyle4, rows = 8:51, cols = 4, gridExpand = FALSE )


# Save workbook
saveWorkbook(wb, file = "Table3_noFormat.xlsx", overwrite = TRUE)

