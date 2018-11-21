############################################################
## Codename - tableA1a_wrangling                            ##
## Data Release - 
## Original Authors - Marios Alexandropoulos              ##
## Orginal Date - November 2018                           ##
## Latest Update Author -   Marios Alexandropoulos        ##
## Latest Update Date -   21/11/18                        ##
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

# select only scotland
smr <-
  smr %>% 
  filter(location_name == "Scotland")



# Create chart smr vs quarter
theme_set(theme_classic())

# Plot
ggplot(smr, aes(x=smr, y=quarter)) + 
  geom_point(col="tomato2", size=3) +   # Draw points
  geom_segment(aes(x=smr, 
                   xend=smr, 
                   y=min(quarter), 
                   yend=max(quarter)), 
               linetype="dashed", 
               size=0.1) +   # Draw dashed lines
  labs(title="Chart 2 - Hospital Standardised Mortality Ratio with Regression Line1; January 2011 - June 2018p", 
       subtitle="SMR Vs quarter", 
       caption="source: mpg") +
  coord_cartesian(ylim = c(0, 1.5)) + 
  coord_flip() 



# Create chart crd_rate vs quarter
theme_set(theme_classic())

# Plot
ggplot(smr, aes(x=crd_rate, y=quarter)) + 
  geom_point(col="blue", size=3) + # Draw points
  geom_line() +
  geom_segment(aes(x=crd_rate, 
                   xend=crd_rate, 
                   y=min(quarter), 
                   yend=max(quarter)), 
               linetype="dashed", 
               size=0.1) +   # Draw dashed lines
  labs(title="Chart 4 - Crude Mortality Rates (%): January 2011 - June 2018p", 
       subtitle="Crude Mortality Rate (%) Vs quarter", 
       caption="source: mpg") +
  coord_cartesian(ylim = c(0, 4)) + 
  coord_flip() 
