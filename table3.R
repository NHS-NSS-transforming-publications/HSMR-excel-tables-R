install.packages("openxlsx")
library("openxlsx")

wb <- loadWorkbook("Table3.xlsx")

sheetName <- "HSMR1"

addWorksheet(wb, sheetName)



style_title <- createStyle(fontSize = 12, textDecoration = "bold", wrapText = FALSE)

writeData(wb, "HSMR1", "Table 3: Impact of the data refresh on the previously published provisional",
          startCol = 2, startRow = 2, rowNames = FALSE)

writeData(wb, "HSMR1", "Standardised Mortality Ratios for all hospitals in Scotland; Oct-Dec 2017",
          startCol = 2, startRow = 3, rowNames = FALSE)

addStyle(wb, "HSMR1", style_title, rows = 2, cols = 2)
addStyle(wb, "HSMR1", style_title, rows = 3, cols = 2)


writeData(wb, "HSMR1", "Oct-Dec 2017",
          startCol = 4, startRow = 5, rowNames = FALSE)

writeData(wb, "HSMR1", "HSMR",
          startCol = 4, startRow = 6, rowNames = FALSE)

writeData(wb, "HSMR1", "as at 15/05/2018",
          startCol = 4, startRow = 7, rowNames = FALSE)

style <- createStyle(fontSize = 12, textDecoration = "bold", wrapText = FALSE)
addStyle(wb, "HSMR1", style, rows = 4:7, cols = 5)


writeData(wb, "HSMR1", "Oct-Dec 2017",
          startCol = 5, startRow = 5, rowNames = FALSE)

writeData(wb, "HSMR1", "Patients",
          startCol = 5, startRow = 6, rowNames = FALSE)

writeData(wb, "HSMR1", "as at 15/05/2018",
          startCol = 5, startRow = 7, rowNames = FALSE)

style <- createStyle(fontSize = 12, textDecoration = "bold", wrapText = FALSE)
addStyle(wb, "HSMR1", style, rows = 4:7, cols = 5)

#####

writeData(wb, "HSMR1", "Oct-Dec 2017",
          startCol = 6, startRow = 5, rowNames = FALSE)

writeData(wb, "HSMR1", "HSMR",
          startCol = 6, startRow = 6, rowNames = FALSE)

writeData(wb, "HSMR1", "as at 14/08/2018",
          startCol = 6, startRow = 7, rowNames = FALSE)

style <- createStyle(fontSize = 12, textDecoration = "bold", wrapText = FALSE)
addStyle(wb, "HSMR1", style, rows = 4:7, cols = 5)


writeData(wb, "HSMR1", "Oct-Dec 2017",
          startCol = 7, startRow = 5, rowNames = FALSE)

writeData(wb, "HSMR1", "Patients",
          startCol = 7, startRow = 6, rowNames = FALSE)

writeData(wb, "HSMR1", "as at 14/08/2018",
          startCol = 7, startRow = 7, rowNames = FALSE)

style <- createStyle(fontSize = 12, textDecoration = "bold", wrapText = FALSE)
addStyle(wb, "HSMR1", style, rows = 4:7, cols = 5)


writeData(wb, "HSMR1", board, startCol = 2, startRow = 8, xy = NULL,
          colNames = TRUE, rowNames = FALSE, headerStyle = NULL,
          borders = c("none"),
          borderColour = getOption("openxlsx.borderColour", "black"),
          borderStyle = getOption("openxlsx.borderStyle", "thin"),
          withFilter = FALSE, keepNA = FALSE, name = NULL, sep = ", ")



writeData(wb, "HSMR1", hospital, startCol = 3, startRow = 9, xy = NULL,
          colNames = TRUE, rowNames = FALSE, headerStyle = NULL,
          borders = c("none"),
          borderColour = getOption("openxlsx.borderColour", "black"),
          borderStyle = getOption("openxlsx.borderStyle", "thin"),
          withFilter = FALSE, keepNA = FALSE, name = NULL, sep = ", ")


# boardStyle <- createStyle(fontSize = 12, fontColour = "#FFFFFF", halign = "center",
#                            fgFill = "#4F81BD", border="TopBottom", borderColour = "#4F81BD")
# addStyle(wb, "HSMR1", boardStyle, rows = 8:52, cols = 2, gridExpand = TRUE)



boardstyle <- createStyle(fontSize = 12, textDecoration = "bold", wrapText = FALSE, halign = "left")
addStyle(wb, "HSMR1", boardstyle, rows = 8:52, cols = 2, gridExpand = TRUE)

hospstyle <- createStyle(fontSize = 12, wrapText = FALSE)
addStyle(wb, "HSMR1", hospstyle, rows = 8:52, cols = 3, gridExpand = TRUE)

setColWidths(wb, "HSMR1", cols = 2:7, widths = "auto")


colourstyle <- createStyle(fontSize = 12, fontName = "Arial",
                           textDecoration = "bold", halign = "left", bgFill = "#FFC7CE", fontColour = "white", border= "TopBottomLeftRight")

addStyle(wb, "HSMR1", colourstyle, rows = 5:7, cols = 2:7, gridExpand = TRUE)

colourstyle1 <- createStyle(fontSize = 10, fontName = "Arial",
                            textDecoration = "bold", halign = "left", bgFill = "#FFC7CE", fontColour = "white", border= "TopBottomLeftRight")

addStyle(wb, "HSMR1", colourstyle1, rows = 7, cols = 2:7, gridExpand = TRUE)



style <- createStyle(fontSize = 12, fontName = "Arial",
                     textDecoration = "bold", halign = "left", bgFill = "#FFC7CE", fontColour = "white", border= "TopBottomLeftRight")

addStyle(wb, "HSMR1", colourstyle, rows = 5:7, cols = 2:7, gridExpand = TRUE)

colourstyle1 <- createStyle(fontSize = 10, fontName = "Arial",
                            textDecoration = "bold", halign = "left", bgFill = "#FFC7CE", fontColour = "white", border= "TopBottomLeftRight")

addStyle(wb, "HSMR1", colourstyle1, rows = 7, cols = 2:7, gridExpand = TRUE)


# source_title <- createStyle(fontSize = 8, textDecoration = "bold", wrapText = TRUE)
# 
# writeData(wb, "HSMR1", "Source: ISD Scotland (SMR01) linked dataset.Reflects the completeness of SMR01 submissions to ISD for individual hospitals as of 12th July 2018.",
#           startCol = 2, startRow = 54, rowNames = FALSE)
# 
# 
# 
# addStyle(wb, "HSMR1", source_title, rows = 54, cols = 2)


borderstyle <- createStyle(fontSize = 12, fontName = "Arial",
                           textDecoration = "bold" , border= "TopBottomLeftRight")

addStyle(wb, "HSMR1", borderstyle, rows = 8:52, cols = 2, gridExpand = FALSE)

borderstyle1 <- createStyle(fontSize = 12, fontName = "Arial",
                            border= "TopBottomLeftRight")

addStyle(wb, "HSMR1", borderstyle1, rows = 8:52, cols = 3, gridExpand = FALSE)


borderstyle2 <- createStyle(fontSize = 12, fontName = "Arial",
                            border= "TopBottomLeftRight")

addStyle(wb, "HSMR1", borderstyle2, rows = 8:52, cols = 4:7, gridExpand = TRUE)



openXL(wb)


















openXL(wb)

board <- c("NHS Ayrshire & Arran", " ", " ","NHS Borders", "" )







board <- c("NHS Ayrshire & Arran",
           " ",
           " ",
           "NHS Borders",
           "",
           "NHS Dumfries & Galloway",
           "",
           "NHS Fife",
           "",
           "NHS Forth Valley",
           "",
           "NHS Grampian",
           "",
           "",
           "NHS Greater Glasgow & Clyde",
           "",
           "",
           "",
           "",
           "NHS Highland",
           "",
           "",
           "",
           "",
           "NHS Lanarkshire",
           "",
           "",
           "",
           "NHS Lothian",
           "",
           "",
           "",
           "National Waiting Times Centre",
           "",
           "NHS Orkney",
           "",
           "NHS Shetland",
           "",
           "NHS Tayside",
           "",
           "",
           "",
           "NHS Western Isles",
           "",
           "Scotland")

hospital <- c(
  "University Hospital Ayr",
  "University Hospital Crosshouse",
  "",
  "Borders General Hospital",
  "",
  "Dumfries & Galloway Royal Infirmary",
  "",
  "Victoria / Queen Margaret / Forth Park",
  "",
  "Forth Valley Royal Hospital",
  "",
  "Aberdeen Royal Infirmary",
  "Dr Gray's Hospital",
  "",
  "Glasgow Royal Infirmary / Stobhill",
  "Inverclyde Royal Hospital",
  "Royal Alexandra / Vale of Leven",
  "Queen Elizabeth University Hospital",
  "",
  "Belford Hospital",
  "Caithness General Hospital",
  "Lorn & Islands Hospital",
  "Raigmore Hospital",
  "",
  "University Hospital Hairmyres",
  "University Hospital Monklands",
  "University Hospital Wishaw",
  "",
  "Royal Infirmary of Edinburgh at Little France",
  "St John's Hospital",
  "Western General Hospital",
  "",
  "Golden Jubilee National Hospital",
  "",
  "Balfour Hospital",
  "",
  "Gilbert Bain Hospital",
  "",
  "Ninewells Hospital",
  "Perth Royal Infirmary",
  "Stracathro Hospital",
  "",
  "Western Isles Hospital",
  ""
)







