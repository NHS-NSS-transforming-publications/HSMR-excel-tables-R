# Using openxlsx
library(readr)
library(openxlsx)
library(dplyr)

#For TableA1a
hsmr_data <- read_csv("smr_data.csv") 
hsmr_data <- hsmr_data %>% filter(location_type == "Scotland"|location_type == "NHS Board")

## load existing workbook
wb <- loadWorkbook("TableA1a_Tina.xlsm")

# write to the workbook
writeData(wb, "smr_data", hsmr_data, startCol = 2, startRow = 1, colNames = TRUE)

# Save workbook     
saveWorkbook(wb, "TableA1a_Tina.xlsm", overwrite = TRUE)

#For TableA2a
long_trend_data <- read_csv("long_term_trends (1).csv") 
long_trend_data <- long_trend_data %>% filter(HB2014 == "Scotland")

## load existing workbook
wb <- loadWorkbook("TableA2a_Tina_test.xlsm")

# write to the workbook
writeData(wb, "long_trend_data", long_trend_data, startCol = 2, startRow = 1, colNames = TRUE)

# Save workbook     
saveWorkbook(wb, "TableA2a_Tina_test.xlsm", overwrite = TRUE)


          
          
          