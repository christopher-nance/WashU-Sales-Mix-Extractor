# WashU-Sales-Mix-Extractor
DRB GSR Sales Mix Extractor (and bonus discount mix extractor!)

Extracts the sales mix for the various wash packages on a retail, monthly and combined level for all sites, including query servers and then totals for the corporation.

Export a GSR as CSV with it split by site and then run the script in the same directory as the csv file. Name the file input.csv. You also need to parse your CSV GSR and change the dictionaries in my code to match the item names in your GSR. 

Report only works with DRB General Sales Reports. Ideally, you can have DRB enable auto exporter for reports, export to dropbox, this goes to the cloud, automation tools like power automate can see the new file, pull it into a microsoft encironment, call an azure python function to process the file and then return the new report back to the server.

Developed by christopher.nance@icloud.com
