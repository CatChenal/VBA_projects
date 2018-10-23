This **VBA Projects Repo** contains code modules for various databases applications I 
developped. 

The code modules belong either to the UI forms for an MS Access front-end or to
general functional, processing steps.  The databases are, obviously not included.



As examplified by module ```MDL_Analysis.bas``` in "PortfolioMonitoringDB - Modules", a common workflow entailed coding
for the automated quarterly report production at a specified date prior to Q end:  

1. Check analyst designated folder for Deal Review documents as per assignments in 
the Analysts table;
2. Email Analyst a missing documents summary;
3. Incorporate additonal data from MS Excel Deal Book
4. Produce a formatted MS Word document for all quarterly Deal reviews.