# -*- coding: utf-8 -*-
"""
This is a python 3.X script designed to access Census Zip Code Business Patterns Data via the Census API, and download the data for a selected list of zip codes into an Excel spreadsheet. This script will go through the following processes:


*   Constructing an API query
*   Cleaning an organizing the data
*   Joining additional attributes
*   Creating Total Economy and Marine Economy dataframes for output
*   Creating analysis tables for output
*   Writing the outputs into an Excel file

# Import Libraries, Set Parameters

The first step is to import the libraries necessary to execute the script, as well as set up the dynamic input parameters.
"""

import pandas as pd
import datetime

StartTime = datetime.datetime.now()
print('Starting at ' + StartTime.strftime("%I:%M:%S %p"))

#################################################################################################
"""
The following items are the inputs to be changed by the user
"""
#This is the list of zip codes requested. These codes will be passed into the url of the API request. If just a single zip code, it will need a comma at the end.
zip_list='54880','55807','55811','55806','55804','55616'

#The year of the data requested. Check the Census Zip Code Business Patterns API to see what years of data are available: https://www.census.gov/data/developers/data-sets/cbp-nonemp-zbp/zbp-api.html
DataYear='2016'

#Set your file path where the output will be saved
OutFile_Loc=r'C:\Data\LocalMarineEconomy\SampleLocation'

#Set the prefix of the Excel file name. The file will append '_MarineEconomy' after the prefix.
OutFile_NamePrefix='Sample'

#################################################################################################
"""
The following items are do not need to be changed, but can be if desired
"""
#Midpoints that will be matched up to employment range codes. These can be customized by the user
Midpoint_List='2.5','7','14.5','34.5','74.5','174.5','374.5','749.5','1000'

#Output file information
OutFile_BaseName='_MarineEconomy.xlsx'
OutFile_Name=OutFile_NamePrefix + OutFile_BaseName
OutFile=OutFile_Loc + '\\' + OutFile_Name
print('Output file: ' + OutFile)

"""# Accessing the Census API

Construct a URL to access the Zip Code Business Patterns data from the Census API

Documentation: https://api.census.gov/data/2016/zbp/variables.html

Example: https://api.census.gov/data/2016/zbp?get=GEO_ID,GEO_TTL,YEAR,NAICS2012_TTL,ESTAB,EMPSZES,EMPSZES_TTL,EMP&for=zipcode:07701,07716&NAICS2012=*

The base url:
https://api.census.gov/data/2016/zbp?get=

The attribute requested:
*   GEO_ID = Unique 14-digit geographic identifier
*   GEO_TTL,YEAR= Descriptive title of the geographic unit
*   NAICS2012_TTL= Descriptive name of the economic industry code
*   ESTAB= Number of business establishments
*   EMPSZES= Code designating the establishment size range by employees
*   EMPSZES_TTL= Descriptive title of the establishment size range by employees
*   zipcode = 5-digit zipcodes data are requested for
*   NAICS2012 = 6-digit industrial codes data are requested for (* is for all codes)
"""

#Loop through creating an API request using each zip code in the list individually.
#This is to get around a Census API limitation
print('Accessing zip code data from the Census API')
df_AllZips = []
for zip in zip_list:
  df='df_' +  zip
  url='https://api.census.gov/data/' + DataYear + '/zbp?get=GEO_TTL,YEAR,NAICS2012_TTL,EMPSZES,EMPSZES_TTL,ESTAB&for=zipcode:'+ zip + '&NAICS2012=*'
  df= pd.read_json(url, orient='values')
  print('Data for zip code ' + zip + ' is being accessed')
  #Set the location of the column headers to be row 0
  df.columns = df.iloc[0]
  #Re-index the data to use the new column headers and drop row 0 from the data values
  df.reindex(df.index.drop(0))
  #Set the dataframe data values to start at row w
  df =df [1:]
  df_AllZips.append(df)
df_AllZips=pd.concat(df_AllZips, ignore_index=True)

print ('All zip codes have been accessed, creating Total Economy table')

#Set the order of the columns
df_AllZips=df_AllZips[['zipcode','GEO_TTL','YEAR','NAICS2012','NAICS2012_TTL','ESTAB','EMPSZES','EMPSZES_TTL']]
#Rename Columns
df_AllZips.rename(columns={'zipcode':'Zipcode','GEO_TTL':'GeoName','YEAR':'Year','NAICS2012':'NAICS'
                           ,'NAICS2012_TTL':'Industry Name','ESTAB':'Establishments','EMPSZES':'Employment Size Class Code'
                           ,'EMPSZES_TTL':'Employment Size Class'}, inplace=True)


"""# Create Data Frames with Additional Attributes

Before we work with the data from the API, we are going to create two dataframes with data we will join in later.

The first table has midpoint values for the emolyment classes. Zip Code Business Patterns provides ranges of employees per establishment in each NAICS code. To be able to do math for an analysis, we need a single value of employees per class, so we take the midpoint of each range.
"""

#Create the data (columns, values)
data={'Employment Size Class Code':['212','220','230','241','242','251','252','254','260']
     ,'Midpoint':Midpoint_List}

#Create a new dataframe using the data above
midpoint_df=pd.DataFrame(data)


"""The second table provides us with the list of NAICS codes that are part of the marine economy as defined by the NOAA Office for Coastal Management. The NAICS codes are grouped into marine sectors for ease of reporting. The crosswalk of NAICS codes th marine sectors can be found here:
https://coast.noaa.gov/data/digitalcoast/pdf/enow-crosswalk-table.pdf
"""

#Create the data (columns, values)
data={'NAICS':['112511','112512','112519','114111','114112','114119','211111','211112','212321','212322','213111','213112','237990','311710','334511','336611','336612',
      '339920','441222','445220','483111','483112','483113','483114','487210','487990','488310','488320','488330','488390','493110','493120','493130','532292',
      '541360','611620','712130','712190','713930','713990','721110','721191','721211','722511','722513','722514','722515'],
      'Marine Sector':['Living Resources','Living Resources','Living Resources','Living Resources','Living Resources','Living Resources','Offshore Mineral Resources','Offshore Mineral Resources'
      ,'Offshore Mineral Resources','Offshore Mineral Resources','Offshore Mineral Resources','Offshore Mineral Resources','Marine Construction','Living Resources','Marine Transportation'
      ,'Ship and Boat Building','Ship and Boat Building','Tourism and Recreation','Tourism and Recreation','Living Resources','Marine Transportation','Marine Transportation','Marine Transportation'
      ,'Marine Transportation','Tourism and Recreation','Tourism and Recreation','Marine Transportation','Marine Transportation','Marine Transportation','Marine Transportation','Marine Transportation'
      ,'Marine Transportation','Marine Transportation','Tourism and Recreation','Offshore Mineral Resources','Tourism and Recreation','Tourism and Recreation','Tourism and Recreation'
      ,'Tourism and Recreation','Tourism and Recreation','Tourism and Recreation','Tourism and Recreation','Tourism and Recreation','Tourism and Recreation','Tourism and Recreation','Tourism and Recreation','Tourism and Recreation']}

#Create a new dataframe using the data above
marinesector_df=pd.DataFrame(data)


"""# Filtering the data

The next  steps will be working with the data from the Census API. First, we want to remove the 'All Establishments' class (EMPSZES code 001). This class just shows the total number of establishments for a NAICS code in a zip code across all employment ranges. This class is not used for the analysis.
"""

#This is where we remove the 'All Establishments' class so we don't double count
RemoveAllEstab=df_AllZips[df_AllZips['Employment Size Class Code']!='001']

#Only keep rows where the values in the NAICS field have a length of 6 characters
RemoveNAICS=RemoveAllEstab[RemoveAllEstab['NAICS'].str.len() == 6]

"""Next, we want to remove any rows where the number of establishments are 0. These are no data points, and removing them will cut down on the amount of data to be analyzed."""

#This is where we filter to only keep rows where establishments are not equal to 0. This cuts down the extraneous data.
TotalEconomy_df=RemoveNAICS[RemoveNAICS['Establishments']!='0']

"""# Create Total Economy Output

Now that we have cleaned up the input data, we are ready to put together the outputs. First, we will create an output for the total economy, which includes everything in each of the zip codes. This output will be used for comparison purposes. The first step is to join the midpoints from the dataframe we created earlier to the API dataframe.
"""

#Here is where we join the midpoints to the total economy data frame, and then set the ESTAB and Midpoint columns to be numeric values
TotalEcon=TotalEconomy_df.join(midpoint_df.set_index('Employment Size Class Code'),on='Employment Size Class Code')

#Set the Estab and Midpoint columns to be numeric so we can do math with the values
TotalEcon[['Establishments','Midpoint']]=TotalEcon[['Establishments','Midpoint']].apply(pd.to_numeric)

"""Next, we will create a new colum called 'EmploymentEstimate', and calculate the value by multiplying the number of establishments with the midpoint."""

#Create the formula
CalcEmp=TotalEcon.Establishments * TotalEcon.Midpoint

#Apply the formula to the new column 'EmploymentEstimate'
TotalEcon['Employment Estimate']=CalcEmp

"""# Create Marine Economy Output

The next step is to take the total economy and filter it down to the industries that are marine dependent.
"""
print('Creating Marine Economy table')

#Filter the total economy by a list of NAICS codes, create the output into a new data frame
MarineEconomy_df=TotalEcon[TotalEcon["NAICS"].isin(['112511','112512','112519','114111','114112','114119','211111','211112','212321','212322','213111'
                                                   ,'213112','237990','311710','334511','336611','336612','339920','441222','445220','483111','483112'
                                                   ,'483113','483114','487210','487990','488310','488320','488330','488390','493110','493120','493130'
                                                   ,'532292','541360','611620','712130','712190','713930','713990','721110','721191','721211','722511'
                                                   ,'722513','722514','722515'])]

"""After creating a new dataframe with just the marine-dependent establishments, we will join in the marine sector titles"""

#Join the marine_df dataframe to the marinesector_df dataframe based on the NAICS2012 column values
Marine_df=MarineEconomy_df.join(marinesector_df.set_index('NAICS'),on='NAICS')

Marine_df.style.hide_index()

"""# Data Analysis

After creating the base data outputs, we will do some initial analysis.First, we will create an output for the total number of establishments and jobs within our list of zip codes.
"""
print('Creating analysis tables')
#Create the dataframe 'TotalEconAnalysis', group the data by the geography attributes, and sum by 'ESTAB' and 'EmploymentEstimate'
TotalEconAnalysis=TotalEcon.groupby(by=['Zipcode','GeoName','Year'])['Establishments','Employment Estimate'].sum().reset_index()

#Create a total for the entire study area
TotalStudyArea_df=TotalEcon.groupby(by=['Year'])['Establishments','Employment Estimate'].sum().reset_index()

#Add back in the Zipcode and GeoName columns
TotalStudyArea_df=TotalStudyArea_df.assign(Zipcode='XXXXX').assign(GeoName='Total for Study Area').assign(Year=DataYear)
TotalStudyArea_df=TotalStudyArea_df[['Zipcode','GeoName','Year','Establishments','Employment Estimate']]

#Append TotalStudyArea_df to TotalEconAnalysis
TotalEconAnalysis=TotalEconAnalysis.append(TotalStudyArea_df, ignore_index=True)
TotalEconAnalysis.Zipcode.apply(str)

#TotalEconAnalysis

"""Next, we will create a series of outputs analyzing the marine economy data. The outputs will include:


*   Marine Economy of  Study Area by Sector
*   Marine Economy of each Zip Code by Sector
*   Marine Economy of Study Area by Industry
*   Marine Economy of each Zip Code by Industry
"""

#Create a total for the entire study area
MarineStudyArea=Marine_df.groupby(by=['Year'])['Establishments','Employment Estimate'].sum().reset_index()

#Add back in the Zipcode and GeoName columns
MarineStudyArea=MarineStudyArea.assign(Zipcode='XXXXX').assign(GeoName='Total for Study Area').assign(Year=DataYear)
MarineStudyArea=MarineStudyArea[['Zipcode','GeoName','Year','Establishments','Employment Estimate']]


#Create a total for the entire study area
MarineStudyAreaZip=Marine_df.groupby(by=['Zipcode','GeoName','Year'])['Establishments','Employment Estimate'].sum().reset_index()


#Append TotalStudyArea_df to TotalEconAnalysis
MarineStudyAreaZip=MarineStudyAreaZip.append(MarineStudyArea, ignore_index=True)
MarineStudyAreaZip.rename(columns={'Establishments':'Marine Establishments','Employment Estimate':'Marine Employment'}, inplace=True)
MarineStudyAreaZip.Zipcode.apply(str)


#Here is where we will join the total economy by zip code and marine economy by zip code tables
TotalEconAnalysis_df=TotalEconAnalysis.merge(MarineStudyAreaZip, on=('Zipcode','GeoName','Year'))
TotalEconAnalysis_df

#Create a new column 'Percent Marine Employment', calculate the values
TotalEconAnalysis_df['Percent Marine Employment']=(TotalEconAnalysis_df['Marine Employment']/TotalEconAnalysis_df['Employment Estimate'])*100
TotalMarineCompare_df=TotalEconAnalysis_df.round(1)

#Marine economy by sector with sums of establishments and employment
MarineSectors=Marine_df.groupby(by=['Year','Marine Sector'])['Establishments','Employment Estimate'].sum().reset_index()
MarineSectors['Average Employment']=MarineSectors['Employment Estimate']/MarineSectors['Establishments']
MarineSectors_df=MarineSectors.round(1)

#Creates a table title for future use in html page
MarineSectorsAnalysis_df=(MarineSectors_df.style.set_caption('Marine Economy by Sector'))

#Marine economy by zip code by sextor with sums of establishments and employment
MarineSectorsZip=Marine_df.groupby(by=['Zipcode','GeoName','Year','Marine Sector'])['Establishments','Employment Estimate'].sum().reset_index()
MarineSectorsZip['Average Employment']=MarineSectorsZip['Employment Estimate']/MarineSectorsZip['Establishments']
MarineSectorsZip_df=MarineSectorsZip.round(1)

#Creates a table title for future use in html page
MarineSectorsZipAnalysis_df=(MarineSectorsZip_df.style.set_caption('Marine Economy by Zip Code and by Sector'))

#Marine economy by industry with sums of establishments and employment
MarineIndustries=Marine_df.groupby(by=['Year','NAICS','Industry Name','Marine Sector'])['Establishments','Employment Estimate'].sum().reset_index()
MarineIndustries['Average Employment']=MarineIndustries['Employment Estimate']/MarineIndustries['Establishments']
MarineIndustries_df=MarineIndustries.round(1)

#Creates a table title for future use in html page
MarineIndustriesAnalysis_df=(MarineIndustries_df.style.set_caption('Marine Economy by Industry'))

#Marine economy by zip code by industry with sums of establishments and employment
MarineIndustriesZip=Marine_df.groupby(by=['Zipcode','GeoName','Year','NAICS','Industry Name','Marine Sector'])['Establishments','Employment Estimate'].sum().reset_index()
MarineIndustriesZip['Average Employment']=MarineIndustriesZip['Employment Estimate']/MarineIndustriesZip['Establishments']
MarineIndustriesZip_df=MarineIndustriesZip.round(1)

#Creates a table title for future use in html page
MarineIndustriesZipAnalysis_df=(MarineIndustriesZip_df.style.set_caption('Marine Economy by Industry'))

"""# Write Outputs To Excel File

Finally, we are going to take the Total Economy dataframe and the Marine Economy dataframe and write them out into separate tabs of an Excel file.
"""
print('Creating the output file')

import pandas.io.formats.excel
pandas.io.formats.excel.header_style = None
#Create the Excel file
writer = pd.ExcelWriter(OutFile, 
                        engine ='xlsxwriter')


#Take the data frames and write them into separate tabs. Horizontal layout of analysis tab.
TotalMarineCompare_df.to_excel(writer, sheet_name='Table1_Analysis', startrow=1, startcol=0,index=False)
MarineSectorsAnalysis_df.to_excel(writer, sheet_name='Table2_Analysis', startrow=1, startcol=0,index=False)
MarineIndustriesAnalysis_df.to_excel(writer, sheet_name='Table3_Analysis', startrow=1, startcol=0,index=False)
MarineSectorsZipAnalysis_df.to_excel(writer, sheet_name='Table4_Analysis', startrow=1, startcol=0,index=False)
MarineIndustriesZipAnalysis_df.to_excel(writer, sheet_name='Table5_Analysis', startrow=1, startcol=0,index=False)
TotalEcon.to_excel(writer, sheet_name='TotalEconomy_Data', startrow=0, startcol=0,index=False)
Marine_df.to_excel(writer, sheet_name='MarineSectors_Data', startrow=0, startcol=0,index=False)

#############################################################################
#Add titles to the analysis tables
workbook  = writer.book
worksheet = writer.sheets['Table1_Analysis']
title1='Table 1 - Comparison of Total Economy and Marine Economy'

#Format the title text
title_format1=workbook.add_format({'bold' : 1
                                  ,'border' : 1
                                  ,'align': 'center'
                                  ,'valign': 'vcenter'})

header_format1=workbook.add_format({ 
                                'bold': True, 
                                'text_wrap': True,
                                'align': 'center',
                                'valign': 'top', 
                                'border': 1}) 

table_format1=workbook.add_format({'text_wrap': True
                                  ,'align': 'center'
                                  ,'valign': 'vcenter'})


#Set column widths for analysis table 1    
worksheet.set_column('A:A',8,table_format1)
worksheet.set_column('B:B',20,table_format1)
worksheet.set_column('C:C',5,table_format1)
worksheet.set_column('D:H',15,table_format1)


for columnnum, columnname in enumerate(list(TotalMarineCompare_df.columns)):
    worksheet.write(1, columnnum, columnname, header_format1)
    
#Merge the cells for the title blocks
worksheet.merge_range('A1:H1',title1,title_format1)

#############################################################################

workbook  = writer.book
worksheet = writer.sheets['Table2_Analysis']
title2='Table 2 - Marine Economy by Sector'


#Format the title text
title_format2=workbook.add_format({'bold' : 1
                                  ,'border' : 1
                                  ,'align': 'center'
                                  ,'valign': 'vcenter'})

header_format2=workbook.add_format({ 
                                'bold': True, 
                                'text_wrap': True,
                                'align': 'center',
                                'valign': 'top', 
                                'border': 1}) 

table_format2=workbook.add_format({'text_wrap': True
                                  ,'align': 'center'
                                  ,'valign': 'vcenter'})
#Set column widths for analysis table 2
worksheet.set_column('A:A',5,table_format2)
worksheet.set_column('B:B',25,table_format2)
worksheet.set_column('C:E',15,table_format2)

for columnnum, columnname in enumerate(list(MarineSectorsAnalysis_df.columns)):
    worksheet.write(1, columnnum, columnname, header_format2)

worksheet.merge_range('A1:E1',title2,title_format2)

#############################################################################

workbook  = writer.book
worksheet = writer.sheets['Table3_Analysis']
title3='Marine Economy by Industry'


#Format the title text
title_format3=workbook.add_format({'bold' : 1
                                  ,'border' : 1
                                  ,'align': 'center'
                                  ,'valign': 'vcenter'})

header_format3=workbook.add_format({ 
                                'bold': True, 
                                'text_wrap': True,
                                'align': 'center',
                                'valign': 'top', 
                                'border': 1}) 

table_format3=workbook.add_format({'text_wrap': True
                                  ,'align': 'center'
                                  ,'valign': 'vcenter'})

#Set column widths for analysis table 3    
worksheet.set_column('A:A',5,table_format3)
worksheet.set_column('B:B',7,table_format3)
worksheet.set_column('C:C',45,table_format3)
worksheet.set_column('D:D',26,table_format3)
worksheet.set_column('E:G',15,table_format3)

for columnnum, columnname in enumerate(list(MarineIndustriesAnalysis_df.columns)):
    worksheet.write(1, columnnum, columnname, header_format3)

worksheet.merge_range('A1:G1',title3,title_format3)


#############################################################################

workbook  = writer.book
worksheet = writer.sheets['Table4_Analysis']
title4='Marine Economy by Zip Code by Sector'


#Format the title text
title_format4=workbook.add_format({'bold' : 1
                                  ,'border' : 1
                                  ,'align': 'center'
                                  ,'valign': 'vcenter'})

header_format4=workbook.add_format({ 
                                'bold': True, 
                                'text_wrap': True,
                                'align': 'center',
                                'valign': 'top', 
                                'border': 1}) 

table_format4=workbook.add_format({'text_wrap': True
                                  ,'align': 'center'
                                  ,'valign': 'vcenter'})

#Set column widths for analysis table 4
worksheet.set_column('A:A',8,table_format4)
worksheet.set_column('B:B',25,table_format4)
worksheet.set_column('C:C',5,table_format4)
worksheet.set_column('D:D',26,table_format4)
worksheet.set_column('E:G',15,table_format4)

for columnnum, columnname in enumerate(list(MarineSectorsZipAnalysis_df.columns)):
    worksheet.write(1, columnnum, columnname, header_format4)
    
worksheet.merge_range('A1:G1',title4,title_format4)

#############################################################################

workbook  = writer.book
worksheet = writer.sheets['Table5_Analysis']
title5='Marine Economy by Zip Code by Industry'


#Format the title text
title_format5=workbook.add_format({'bold' : 1
                                  ,'border' : 1
                                  ,'align': 'center'
                                  ,'valign': 'vcenter'})

table_format5=workbook.add_format({'text_wrap': True
                                  ,'align': 'center'
                                  ,'valign': 'vcenter'})

header_format5=workbook.add_format({ 
                                'bold': True, 
                                'text_wrap': True,
                                'align': 'center',
                                'valign': 'top', 
                                'border': 1}) 


#Set column widths for analysis table 5
worksheet.set_column('A:A',8,table_format5)
worksheet.set_column('B:AB',20,table_format5)
worksheet.set_column('C:C',5,table_format5)
worksheet.set_column('D:D',7,table_format5)
worksheet.set_column('E:E',45,table_format5)
worksheet.set_column('F:F',26,table_format5)
worksheet.set_column('G:I',15,table_format5)

for columnnum, columnname in enumerate(list(MarineIndustriesZipAnalysis_df.columns)):
    worksheet.write(1, columnnum, columnname, header_format5)
    
worksheet.merge_range('A1:I1',title5,title_format5)
#############################################################################

#Save the Excel file
writer.save()

print('Your file is ready!')
EndTime = datetime.datetime.now()
print('Ended at ' + EndTime.strftime("%I:%M:%S %p"))
