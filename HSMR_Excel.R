# Using openxlsx
library(readr)
library(openxlsx)
library(dplyr)


# For TableA1a----
hsmr_data <- read_csv("smr_data.csv") 
hsmr_data_A1a <- hsmr_data %>% filter(location_type == "Scotland"|location_type == "NHS Board")

## load existing workbook
wb <- loadWorkbook("TableA1a_Tina.xlsm")

# write to the workbook
writeData(wb, "smr_data", hsmr_data_A1a, startCol = 2, startRow = 1, colNames = TRUE)

# Save workbook     
saveWorkbook(wb, "TableA1a_Tina.xlsm", overwrite = TRUE)


# For Table2----
hsmr_data_2 <- hsmr_data %>% filter(location_type == "Scotland"|location_type == "hospital")

## load existing workbook
wb <- loadWorkbook("Table2.xlsx")

# write to the workbook
writeData(wb, "smr_data", hsmr_data_2, startCol = 2, startRow = 1, colNames = TRUE)

# Save workbook     
saveWorkbook(wb, "Table2.xlsx", overwrite = TRUE)


# For Table3----
hsmr_data_3 <- hsmr_data %>% filter(location_type == "Scotland"|location_type == "hospital")

hsmr_data_prev <- read_csv("smr_data_prev.csv") 
hsmr_data_prev <- hsmr_data_prev %>% filter(location_type == "Scotland"|location_type == "hospital")

## load existing workbook
wb <- loadWorkbook("Table3.xlsx")

# write to the workbook
writeData(wb, "smr_data", hsmr_data_3, startCol = 2, startRow = 1, colNames = TRUE)
writeData(wb, "smr_data_prev", hsmr_data_prev, startCol = 2, startRow = 1, colNames = TRUE)

# Save workbook     
saveWorkbook(wb, "Table3.xlsx", overwrite = TRUE)


# For TableA2a----
long_trend_data <- read_csv("long_term_trends (1).csv") 
long_trend_data_A2a <- long_trend_data %>% filter(HB2014 == "Scotland")

## load existing workbook
wb <- loadWorkbook("TableA2a_Tina.xlsx")

# write to the workbook
writeData(wb, "long_trend_data", long_trend_data_A2a, startCol = 2, startRow = 1, colNames = TRUE)

# Save workbook     
saveWorkbook(wb, "TableA2a_Tina.xlsx", overwrite = TRUE)


          
          
          