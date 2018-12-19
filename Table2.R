############################################################
## Codename - table2_wrangling                            ##
## Data Release - 
## Original Authors - Marios Alexandropoulos              ##
## Orginal Date - December 2018                           ##
## Latest Update Author -   Marios Alexandropoulos        ##
## Latest Update Date -   19/12/18                        ##
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
## Description -  Template for producing Table 3:         ## 
## Table 2: Percentage Change1 in Standardised Mortality  ##
##  Ratios (regression line)                              ##
## for hospitals in Scotland between                      ##
##  Jan-Mar 14 and Apr-Jun 18p                            ##
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
library(r2excel) # used only for adding hyperlink

# Create a workbook
wb <- loadWorkbook("Table2.xlsx")

sheetName <- "Percentage"

# Add sheet
addWorksheet(wb, sheetName)


# Add and format title

style_title <- createStyle(fontSize = 12, textDecoration = "bold", wrapText = FALSE)

writeData(wb, "Percentage", "Table 2: Percentage Change^1 in Standardised Mortality Ratios (regression line) ",
          startCol = 2, startRow = 2, rowNames = FALSE)

writeData(wb, "Percentage", "for hospitals in Scotland between Jan-Mar 14 and Apr-Jun 18p",
          startCol = 2, startRow = 3, rowNames = FALSE)

addStyle(wb, "Percentage", style_title, rows = 2, cols = 2)
addStyle(wb, "Percentage", style_title, rows = 3, cols = 2)


# Add HB and location columns
loc_code <- c("A210H", "A111H", " ", "B120H", " " ,"Y146H", " ", "F704H", " ", "V217H", " ", "N101H", "N411H", " ", "G107H", "C313H", "C418H",
              "G405H", " ", "H212H", "H103H", "C121H", "H202H", " ", "L302H", "L106H", "L308H", " ", "S314H", "S308H", "S116H", " ", "D102H",
              " ", "R101H", " ", "Z102H", " ", "T101H", "T202H", "T312H", " ", "W107H", "Scot")

HB_loc <- c("NHS Ayrshire & Arran" ,"University Hospital Ayr", "University Hospital Crosshouse", "NHS Borders", "Borders General Hospital",
            "NHS Dumfries & Galloway", "Dumfries & Galloway Royal Infirmary", "NHS Fife", "Victoria / Queen Margaret / Forth Park", "NHS Forth Valley",
            "Forth Valley Royal Hospital", "NHS Grampian", "Aberdeen Royal Infirmary", "Dr Gray's Hospital", "NHS Greater Glasgow & Clyde", "Glasgow Royal Infirmary / Stobhill",
            "Inverclyde Royal Hospital", "Royal Alexandra / Vale of Leven", "Queen Elizabeth University Hospital", "NHS Highland",
            "Belford Hospital", "Caithness General Hospital", "Lorn & Islands Hospital",
            "Raigmore Hospital", "NHS Lanarkshire", "University Hospital Hairmyres", "University Hospital Monklands", "University Hospital Wishaw",
            "NHS Lothian", "Royal Infirmary of Edinburgh at Little France", "St John's Hospital", "Western General Hospital", "National Waiting Times Centre",
            "Golden Jubilee National Hospital", "NHS Orkney", "Balfour Hospital", "NHS Shetland", "Gilbert Bain Hospital", "NHS Tayside",
            "Ninewells Hospital", "Perth Royal Infirmary", "Stracathro Hospital",
            "NHS Western Isles", "Western Isles Hospital", "Scotland")



writeData(wb, "Percentage", loc_code, startCol = 1, startRow = 8, xy = NULL,
          colNames = TRUE, rowNames = FALSE, headerStyle = NULL,
          borders = c("none"),
          borderColour = getOption("openxlsx.borderColour", "white"),
          borderStyle = getOption("openxlsx.borderStyle", "thin"),
          withFilter = FALSE, keepNA = FALSE)



writeData(wb, "Percentage", HB_loc, startCol = 2, startRow = 7, xy = NULL,
          colNames = TRUE, rowNames = FALSE, headerStyle = NULL,
          borders = c("all"),
          borderColour = getOption("openxlsx.borderColour", "black"),
          borderStyle = getOption("openxlsx.borderStyle", "thin"),
          withFilter = FALSE, keepNA = FALSE)


# Format HB and location columns
loc_codestyle <- createStyle(fontSize = 5, wrapText = FALSE, halign = "left",fontColour = "white")
addStyle(wb, "Percentage", loc_codestyle, rows = 7:51, cols = 1, gridExpand = TRUE)


HB_locstyle <- createStyle(fontSize = 12, wrapText = FALSE, fontName = "Arial", border= "TopBottomLeftRight")
addStyle(wb, "Percentage", HB_locstyle, rows = 7:51, cols = 2, gridExpand = TRUE)

boardstyle <- createStyle(fontSize = 12, textDecoration = "bold", wrapText = FALSE, halign = "left", fontName = "Arial", border= "TopBottomLeftRight")
addStyle(wb, "Percentage", boardstyle, rows = c(7,10,12,14,16,18,21,26,31,35,39,41,43,45,49,51), cols = 2, gridExpand = TRUE)




# Add and format table's names


writeData(wb, "Percentage", "        Jan-Mar 2014",
          startCol = 3, startRow = 5, rowNames = FALSE)


mergeCells(wb, "Percentage", rows = 5, cols = c(3,4))

style <- createStyle(fontSize = 12, textDecoration = "bold", wrapText = FALSE)
addStyle(wb, "Percentage", style, rows = 5, cols = c(3,4) )

writeData(wb, "Percentage", "        Apr-Jun 2018p",
          startCol = 5, startRow = 5, rowNames = FALSE)


mergeCells(wb, "Percentage", rows = 5, cols = c(5,6))

style1 <- createStyle(fontSize = 12, textDecoration = "bold", wrapText = FALSE)
addStyle(wb, "Percentage", style1, rows = 5, cols = c(5,6))


writeData(wb, "Percentage", "% Change",
          startCol = 7, startRow = 5, rowNames = FALSE)

removeCellMerge(wb, "Percentage", cols = c(7,8), rows = 5)


style2 <- createStyle(fontSize = 12, textDecoration = "bold", wrapText = FALSE)
addStyle(wb, "Percentage", style2, rows = 5, cols = 7)
                                                      
                                                    
writeData(wb, "Percentage", "SMR",
          startCol = 3, startRow = 6, rowNames = FALSE)


style3 <- createStyle(fontSize = 12, textDecoration = "bold", wrapText = FALSE)
addStyle(wb, "Percentage", style3, rows = 6, cols = 3)


writeData(wb, "Percentage", "Regression line",
          startCol = 4, startRow = 6, rowNames = FALSE)


style4 <- createStyle(fontSize = 12, wrapText = TRUE)
addStyle(wb, "Percentage", style4, rows = 6, cols = 4)


writeData(wb, "Percentage", "SMR",
          startCol = 5, startRow = 6, rowNames = FALSE)


style5 <- createStyle(fontSize = 12, textDecoration = "bold",wrapText = FALSE)
addStyle(wb, "Percentage", style5, rows = 6, cols = 5)

writeData(wb, "Percentage", "Regression line",
          startCol = 6, startRow = 6, rowNames = FALSE)


style6 <- createStyle(fontSize = 12, wrapText = TRUE)
addStyle(wb, "Percentage", style6, rows = 6, cols = 6)



writeData(wb, "Percentage", "calculated from regression line1",
          startCol = 7, startRow = 6, rowNames = FALSE)


style7 <- createStyle(fontSize = 10, wrapText = TRUE)
addStyle(wb, "Percentage", style7, rows = 6, cols = 7)


# Set columns widths
setColWidths(wb, "Percentage", cols = c(1,2,3,4,5,6,7),ignoreMergedCells = TRUE, widths = c(4,43,12,12,12,12,13 ))

# Add and format source 
source_title <- createStyle(fontSize = 9, fontName = "Arial", wrapText = FALSE)

writeData(wb, "Percentage", "p = Provisional",
          startCol = 2, startRow = 53, rowNames = FALSE)

addStyle(wb, "Percentage", source_title, rows = 53, cols = 2)

source_title1 <- createStyle(fontSize = 9, fontName = "Arial", wrapText = FALSE)

writeData(wb, "Percentage", "Source: ISD Scotland (SMR01) linked dataset. Reflects the completeness of SMR01 submissions to ISD for individual hospitals as of 12th October 2018.",
          startCol = 2, startRow = 54, rowNames = FALSE)

addStyle(wb, "Percentage", source_title1, rows = 54, cols = 2)


source_title2 <- createStyle(fontSize = 9, fontName = "Arial", wrapText = FALSE)

writeData(wb, "Percentage", "1. The percentage change is measured against the difference between the regression line values of January to March 2014 (first after baseline) and the latest quarter.",
          startCol = 2, startRow = 55, rowNames = FALSE)

addStyle(wb, "Percentage", source_title2, rows = 55, cols = 2)


source_title3 <- createStyle(fontSize = 9, fontName = "Arial", wrapText = FALSE)

writeData(wb, "Percentage", "Please refer to the following paper for further information on how the regression line is calculated.",
          startCol = 2, startRow = 56, rowNames = FALSE)

addStyle(wb, "Percentage", source_title3, rows = 56, cols = 2)

# r2excel package is required but it doesn't work for older version of R
#xlsx.addHyperlink(wb, "Percentage", "https://www.isdscotland.org/Health-Topics/Quality-Indicators/HSMR/FAQ/_docs/HSMR_regression_line.pdf", "How Regression line is calculated", fontSize=9)

openXL(wb)

#saveWorkbook(wb, file = "Table2_Format.xlsx", overwrite = TRUE)


