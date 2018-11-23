############################################################
## Codename - tableA1a_wrangling                            ##
## Data Release - 
## Original Authors - Marios Alexandropoulos              ##
## Orginal Date - November 2018                           ##
## Latest Update Author -   Marios Alexandropoulos        ##
## Latest Update Date -   22/11/18                        ##
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
## Description -  Template for producing Table A1a: 
## Quarterly Hospital Standardised Mortality Ratios                                 ## 
##                                                        ##
## Approximate run time:                                  ## 
############################################################



################################## 
##################################
### SECTION 1 - HOUSE KEEPING ----
##################################
##################################


library(magrittr)
library(dplyr)
library(ggplot2)
library(scales)

# Read in smr data
smr <- read.csv("smr_data.csv", stringsAsFactors = FALSE)
str(smr)
head(smr)


# Create quarter_name
smr <- smr %>% 
  mutate(quarter_name = ifelse(quarter == 1, "Jan-Mar2011",
                               ifelse(quarter == 2, "Apr-Jun2011", 
                                      ifelse(quarter == 3, "Jul-Sep2011",
                                             ifelse(quarter == 4, "Oct-Dec2011",
                                                    ifelse(quarter == 5, "Jan-Mar2012", 
                                                           ifelse(quarter == 6, "Apr-Jun2012",
                                                                  ifelse(quarter == 7, "Jul-Sep2012",
                                                                         ifelse(quarter == 8, "Oct-Dec2012", 
                                                                                ifelse(quarter == 9, "Jan-Mar2013",
                                                                                       ifelse(quarter == 10, "Apr-Jun2013",
                                                                                              ifelse(quarter == 11, "Jul-Sep2013", 
                                                                                                     ifelse(quarter == 12, "Oct-Dec2013", 
                                                                                                            ifelse(quarter == 13, "Jan-Mar2014", 
                                                                                                                   ifelse(quarter == 14, "Apr-Jun2014",
                                                                                                                          ifelse(quarter == 15, "Jul-Sep2014",
                                                                                                                                 ifelse(quarter == 16, "Oct-Dec2014",
                                                                                                                                        ifelse(quarter == 17, "Jan-Mar2015",
                                                                                                                                               ifelse(quarter == 18, "Apr-Jun2015",
                                                                                                                                                      ifelse(quarter == 19, "Jul-Sep2015",
                                                                                                                                                             ifelse(quarter == 20, "Oct-Dec2015",
                                                                                                                                                                    ifelse(quarter == 21, "Jan-Mar2016",
                                                                                                                                                                           ifelse(quarter == 22, "Apr-Jun2016",
                                                                                                                                                                                  ifelse(quarter == 23, "Jul-Sep2016",
                                                                                                                                                                                         ifelse(quarter == 24, "Oct-Dec2016",
                                                                                                                                                                                                ifelse(quarter == 25, "Jan-Mar2017",
                                                                                                                                                                                                       ifelse(quarter == 26, "Apr-Jun2017",
                                                                                                                                                                                                              ifelse(quarter == 27, "Jul-Sep2017",
                                                                                                                                                                                                                     ifelse(quarter == 28, "Oct-Dec2017",
                                                                                                                                                                                                                            ifelse(quarter == 29, "Jan-Mar2018",'NULL'))))))))))))))))))))))))))))))

# Check all quarters have matched properly
quarter_check <- table(smr$quarter_name, exclude = NULL)
quarter_check                                                                                                                                                                                                                                  


# select only scotland
smr_scot <-
  smr %>% 
  filter(location_name == "Scotland")

# Create chart smr vs quarter
theme_set(theme_classic())

# Plot
ggplot(smr_scot, aes(x=smr, y=quarter_name)) + 
  geom_point(col="blue", size=3) +   # Draw points
  geom_segment(aes(x=smr, 
                   xend=smr, 
                   y=min(quarter_name), 
                   yend=max(quarter_name)), 
               linetype="dashed", 
               size=0.1) +   # Draw dashed lines
  labs(title="Quarterly Hospital Standardised Mortality Ratios in Scotland
       Mortality within 30-days of admission", 
       subtitle="Chart 2 - Hospital Standardised Mortality Ratio with Regression Line1; January 2011 - June 2018p", 
       caption="
       
       source: ISD Scotland (SMR01) linked dataset. Reflects the completeness of SMR01 submissions to ISD for individual hospitals as of 12th October 2018.
       P = provisional (see main report Appendix A1 - Data Quality and Timeliness).
       1. A regression line through data points from January to March 2014, to the current HSMR is used to smooth out seasonal variations in HSMR and to monitor long term change.") +
  coord_cartesian(ylim = c(0, 1.5)) + 
  coord_flip() +
  theme(axis.text.x = element_text(angle = 90, hjust = 1))



# Create chart crd_rate vs quarter
theme_set(theme_classic())



# Plot
ggplot(smr_scot, aes(x=crd_rate, y=quarter_name)) + 
  geom_point(col="blue", size=3) + # Draw points
  geom_line() +
  geom_segment(aes(x=crd_rate, 
                   xend=crd_rate, 
                   y=min(quarter_name), 
                   yend=max(quarter_name)), 
               linetype="dashed", 
               size=0.1) +   # Draw dashed lines
  labs(title="Quarterly Hospital Standardised Mortality Ratios in Scotland
       Mortality within 30-days of admission", 
       subtitle="Chart 4 - Crude Mortality Rates (%): January 2011 - June 2018p", 
       caption="
       
       Source: ISD Scotland (SMR01) linked dataset. Reflects the completeness of SMR01 submissions to ISD for individual hospitals as of 12th October 2018.
       P = provisional (see main report Appendix A1 - Data Quality and Timeliness).") +
  coord_cartesian(ylim = c(0, 4)) + 
  coord_flip() +
  theme(axis.text.x = element_text(angle = 90, hjust = 1)) 



# # creat workbook and sheet
wb_A1a <- createWorkbook(type="xlsx")
sheet <- createSheet(wb_A1a, sheetName = "Chart_4")

# Add header
xlsx.addHeader(wb_A1a, sheet, "Quarterly Hospital Standardised Mortality Ratios in Scotland Mortality within 30-days of admission", level= 2)
xlsx.addHeader(wb_A1a, sheet, "Chart 4 - Crude Mortality Rates (%): January 2011 - June 2018p", level= 3, color = "darkgrey")

xlsx.addLineBreak(sheet, 1)

# Plot function
plotFunction<-function(){
  p_crd <- ggplot(smr_scot, aes(x=crd_rate, y=quarter_name)) + 
    geom_point(col="blue", size=3) + # Draw points
    geom_line() +
    geom_segment(aes(x=crd_rate, 
                     xend=crd_rate, 
                     y=min(quarter_name), 
                     yend=max(quarter_name)), 
                 linetype="dashed", 
                 size=0.1) +   # Draw dashed lines
    labs(title="", 
         subtitle="", 
         caption="
         
         Source: ISD Scotland (SMR01) linked dataset. Reflects the completeness of SMR01 submissions to ISD for individual hospitals as of 12th October 2018.
         P = provisional (see main report Appendix A1 - Data Quality and Timeliness).") +
    coord_cartesian(ylim = c(0, 4)) + 
    coord_flip() +
    theme(axis.text.x = element_text(angle = 90, hjust = 1))
  print(p_crd)
}

# Add plot in the excel sheet
xlsx.addPlot(wb_A1a, sheet, plotFunction())

# save the workbook to an Excel file.
saveWorkbook(wb_A1a, "table_A1a.xlsx")

xlsx.openFile("table_A1a.xlsx")# view the file
