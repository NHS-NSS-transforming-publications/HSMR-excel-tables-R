############################################################
## Codename - table3_wrangling                            ##
## Data Release - 
## Original Authors - Marios Alexandropoulos              ##
## Orginal Date - November 2018                           ##
## Latest Update Author -   Marios Alexandropoulos        ##
## Latest Update Date -   08/11/18                        ##
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
## hospitals in Scotland                                  ##
##                                                        ## 
##                                                        ##
## Approximate run time:                                  ## 
############################################################



################################## 
##################################
### SECTION 1 - HOUSE KEEPING ----
##################################
##################################


# load Packages
install.packages("devtools")
devtools::install_github("kassambara/r2excel")
library(r2excel)
install.packages("xlsx")

install.packages("tidyverse")
library(readr)

library(dplyr)
library(tidyr)
library(magrittr) # for %<>% operator
library(tibble)
library("r2excel")


# Create an Excel workbook. 
# Both .xls and .xlsx file formats can be used.
filename <- "r2excel-example1.xlsx"
wb <- createWorkbook(type="xlsx")
# Create a sheet in that workbook to contain the data table
sheet <- createSheet(wb, sheetName = "HSMR")


# Add header
xlsx.addHeader(wb, sheet, value="Table 3: Impact of the data refresh on the previously published provisional",
               level=4, color="black", underline=1)

xlsx.addHeader(wb, sheet, value="Standardised Mortality Ratios for all hospitals in Scotland; Oct-Dec 2017",
               level=4, color="black", underline=1)

# for adding a line break. This is useful to skip lines between two data frames.
xlsx.addLineBreak(sheet, 1)

# Add paragraph : Author
par=paste("Source: ISD Scotland (SMR01) linked dataset. 
             Reflects the completeness of SMR01 submissions 
             to ISD for individual hospitals as of 12th July 2018.", sep="")
xlsx.addParagraph(wb, sheet,value=par, isItalic=TRUE, colSpan=5, 
                  rowSpan=4, fontColor="darkgray", fontSize=14)
xlsx.addLineBreak(sheet, 3)

# Add Hyperlink
xlsx.addHyperlink(wb,sheet, address= "http://www.isdscotland.org/Health-Topics/Quality-Indicators/HSMR/", "HSMR") 



  
# Read in table 3 
tab <- read.xlsx2(file ="Table3.xlsx" , sheetIndex = 1, startRow =5 )  
head(tab)
str(tab)

# Create as tibble
tab <- as.tibble(tab)

# Convert data from factor to numeric
tab$Oct.Dec.2017.HSMR <- as.numeric(as.character(tab$Oct.Dec.2017.HSMR))
tab$Oct.Dec.2017.Patients<- as.numeric(as.character(tab$Oct.Dec.2017.Patients))
tab$Oct.Dec.2017.HSMR.1 <- as.numeric(as.character(tab$Oct.Dec.2017.HSMR.1))
tab$Oct.Dec.2017.Patients.1<- as.numeric(as.character(tab$Oct.Dec.2017.Patients.1))



#tab      <- tab %>% 
#  rename(HB = X.) %>% 
#  rename(Hospital = X..1) %>%  
#  transmute(Oct.Dec.2017.Patients = as.numeric(as.character(Oct.Dec.2017.Patients))) %>%  
#  transmute(Oct.Dec.2017.HSMR = as.numeric(as.character(Oct.Dec.2017.HSMR))) %>%
#  transmute(Oct.Dec.2017.HSMR.1 = as.numeric(as.character(Oct.Dec.2017.HSMR.1))) %>%
#  transmute(Oct.Dec.2017.Patients.1 = as.numeric(as.character(Oct.Dec.2017.Patients.1))) %>%
#  na.omit() 
   

names(tab)[1:2] <- c("NHS_HB", "NHS Hospital") 

# Delete blank rows
tab <- tab[-c(47:53), ]
tab <- tab[-c(1),]




# Add table : add a data frame
xlsx.addTable(wb, sheet, tab, startCol=2)
xlsx.addLineBreak(sheet, 2)

# save the workbook to an Excel file
saveWorkbook(wb, filename)


