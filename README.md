# Local Marine Economy Data Extraction and Analysis

This is a python 3.X script designed to access Census Zip Code Business Patterns Data via the Census API, download the data for a selected list of zip codes, organize and group the data based on a list of NAICS codes, and export the information into a formatted Excel spreadsheet with multiple tabs. This script will go through the following processes:

* Constructing an API query
* Cleaning an organizing the data
* Joining additional attributes
* Creating Total Economy and Marine Economy dataframes for output
* Creating analysis tables for output
* Writing the outputs into an Excel file

The information needed to run this file are:
* A single zip code or list of zip codes
* The year the data will be run for
* An output file prefix
* A file path for the output file

The data produced in the script are used in the [Estimating the Local Marine Economy training](https://coast.noaa.gov/digitalcoast/training/marine-economy.html) delivered by the NOAA Office for Coastal Management.

For additional information, contact:  
Gabe Sataloff  
CSS at the NOAA Office for Coastal Management  
gabe.sataloff@noaa.gov

# NOAA Open Source Disclaimer

This repository is a scientific product and is not official communication of the National Oceanic and Atmospheric Administration, or the United States Department of Commerce. All NOAA GitHub project code is provided on an as is basis and the user assumes responsibility for its use. Any claims against the Department of Commerce or Department of Commerce bureaus stemming from the use of this GitHub project will be governed by all applicable Federal law. Any reference to specific commercial products, processes, or services by service mark, trademark, manufacturer, or otherwise, does not constitute or imply their endorsement, recommendation or favoring by the Department of Commerce. The Department of Commerce seal and logo, or the seal and logo of a DOC bureau, shall not be used in any manner to imply endorsement of any commercial product or activity by DOC or the United States Government.

# License

Software code created by U.S. Government employees is not subject to copyright in the United States (17 U.S.C. 105). The United States/Department of Commerce reserve all rights to seek and obtain copyright protection in countries other than the United States for Software authored in its entirety by the Department of Commerce. To this end, the Department of Commerce hereby grants to Recipient a royalty-free, nonexclusive license to use, copy, and create derivative works of the Software outside of the United States.
