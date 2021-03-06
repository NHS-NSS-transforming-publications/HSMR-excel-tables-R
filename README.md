# HSMR-excel-tables-R
Template for producing publication tables in R so that the publication process can be streamlined

Notes:

1. David should add a column called "quarter_name" (e.g. Jan-Mar 2011) into smr_data.csv and long_term_trends(1).csv, next to "quarter". 

2. For Table 3, currently we are using the same dataset for spreadsheets "smr_data" and "smr_data_prev", as there is no data yet from next publication. So the results for spreadsheet "HSMR" columns F and G are not referring to "Oct-Dec 2017", but once the data from next publication got populated, it should refer to the correct data as the formulae have been set up correctly.

3. I noticed the "location" in smr_data.csv is still using old GSS codes for Fife and Tayside. But you should use new codes due to council area boundary change. 

4. David should modify the R codes to make the previous smr_data.csv as smr_data_prev.csv because Table 3 needs both of the datasets. 

5. David should make long_term_trends(1).csv always have quarter 1 to 40 as rolling quarters. 

6. We found for "Dumfries & Galloway Royal Infirmary" there are only data for Oct-Dec 2017 and Jan-Mar 2018. It is because it is a new hospital opening from 1st Jan 2017. So we can't get the data for previous quarters (e.g. you can see no data returned in Table 2).

