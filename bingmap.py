## README ##
## This script takes an input file of from and to addresses to find drive time and distance via Bing Maps.
## It then calculates total and average drive time and distance, summarized by Region, Market, and FSM.
## Before using this script, please install the following packages: requests, pandas, datetime, and xlsxwriter.(IE: pip install requests)
## The csv, json, and os packages should have already been downloaded when installing Python 3.
## Bing Maps is free if you consume fewer than 125,000 transactions in a 12-month period.
## Get a new Bing Maps Key or check your usage here: https://www.bingmapsportal.com/Application#
##
## Additional Information
## Bing Rest Routes API documentation: https://msdn.microsoft.com/en-us/library/mt270292.aspx
## Use this url format when incorporating waypoints:
##  routeUrl = "http://dev.virtualearth.net/REST/V1/Routes/Driving?o=json&wp.0=" + str(wp0) + "&vwp.1=" + str(wp1) + "&wp.2=" + str(wp2) + "&distanceUnit=Mile&key==" + bingMapsKey
## Use this url format for routes from points A to B:
##  http://dev.virtualearth.net/REST/V1/Routes/Driving?o=json&wp.0=3427%Hopi%Point,%Lawrenceville,%GA&wp.1=2201%HENDERSON%MILL%ROAD%NORTHEAST,%ATLANTA,%GA&distanceUnit=Mile&key=j8ROx37qaPdF6DV0GkMD~UmHYzVMvhzv6p2z2NZLyzQ~AscsmLIAN6JAxrv4UmtfNo8mUjnUb1kMDhS1tlcGcmNWVjQgHczpkrTRD3oas6vM

## Count of total visits for region, market, and FSM do not include routes that gave errors or list "0" for travel distance. Counts only include
## successful route pulls with distance greater than "0" (any routes with from/to address as same address do not add to the count)

##  Additions/Changes 
##  - remove home addresses for first & last visits from input file
##  (Done) - if Bing give an error in pulling the directions, or if from/to distance is 0, remove from the count summary
##  (Done) - if Bing give an error in pulling the directions, or if from/to distance is 0, remove from the count when finding averages
##  - add manipulation of input file directly from FMO output (no extra entries for home addresses)
##  - fix time output fields, currently in minutes


import requests
import csv
import json
import pandas
from pandas import DataFrame
import datetime
import os
import xlsxwriter

#set folder path to where the files should go
os.chdir('C:/Users/ahanway/Desktop')

# Bing Maps Key 
bingMapsKey = "j8ROx37qaPdF6DV0GkMD~UmHYzVMvhzv6p2z2NZLyzQ~AscsmLIAN6JAxrv4UmtfNo8mUjnUb1kMDhS1tlcGcmNWVjQgHczpkrTRD3oas6vM"

##creates output file as xlsx with worksheets for summary and detail
writer = pandas.ExcelWriter('BingMapsOutput.xlsx', engine='xlsxwriter')
workbook  = writer.book

##creates output as csv
outputfile = open('BingMapsOutput.csv', 'w', newline='')
csvwriter = csv.writer(outputfile)
header = "Distance (Miles)", "Time (Minutes from seconds)", "Time (Seconds)"
csvwriter.writerow(header)

payload = {'type': 'adminVM', 'pageSize': '100', 'filter': 'status==POWERED_ON'}

#opening the input file, iterating through each row to create urls with from/to address fields, 
#pulling data from json requests and writing to the output file
with open("C:/Users/ahanway/Desktop/Store Visit Mapping/HE Store Visit Drive Time4.csv") as csvfile: 
	rowreader = csv.reader(csvfile)
	totalrows = -1
	for row in rowreader:
		totalrows += 1		
	
with open("C:/Users/ahanway/Desktop/Store Visit Mapping/HE Store Visit Drive Time4.csv") as csvfile: 
	reader = csv.DictReader(csvfile)
	counter = 0
	errorcount = 0
	totalrows = int(totalrows)
	if counter <= totalrows:
		for row in reader:
			try:
				counter += 1
				print("Processing Input Line: " + str(counter) + " of " + str(totalrows))
				wp0 = row['FromAddress']
				wp1 = row['ToAddress']
				routeUrl = "http://dev.virtualearth.net/REST/V1/Routes/Driving?o=json&wp.0=" + str(wp0) + "&wp.1=" + str(wp1) + "&distanceUnit=Mile&key=" + bingMapsKey
				page = requests.get((routeUrl), params=payload, verify=False).json()
				distance = (page['resourceSets'][0]['resources'][0]['travelDistance'])
				traveltime = (page['resourceSets'][0]['resources'][0]['travelDurationTraffic'])
				traveltime_seconds = (page['resourceSets'][0]['resources'][0]['travelDurationTraffic'])
				#traveltime = str(datetime.timedelta(seconds = traveltime))
				traveltime = traveltime_seconds / 60
			except:
				print("Error processing")
				errorcount +=1
				distance = "0"
				traveltime = "0"
				traveltime_seconds = "0"
			with open('BingMapsOutput.csv') as f:
				values = distance, traveltime, traveltime_seconds
				csvwriter.writerow(values)
		
outputfile.close()
print("Finished Processing")
print("Creating Files...")

#append output file to inputfile
left = pandas.read_csv('C:/Users/ahanway/Desktop/Store Visit Mapping/HE Store Visit Drive Time4.csv', encoding = "ISO-8859-1")
right = pandas.read_csv('BingMapsOutput.csv', encoding = "ISO-8859-1")
result = left.join(right, how="inner")
result.to_csv('BingMapsOutputDetail.csv', index=False)

######find totals and averages######
#creates new worksheet summarizing fsm and region time/distance/averages

#fsm/market summary of total time, total distance, and total visits	
summaryoutput = open('BingMapsFSMSummary.csv', 'w', newline='')
df = pandas.read_csv('BingMapsOutputDetail.csv', encoding = "ISO-8859-1")
Total_Distance = df.groupby('FSM')['Distance (Miles)'].sum()
Total_Time = df.groupby('FSM')['Time (Seconds)'].sum()
Total_Time = Total_Time / 60
#Total_Time = str(datetime.timedelta(seconds = Total_Time))
##Total_Visits = df.groupby('FSM')['FromAddress'].count()
Total_Visits_df = df[(df['Distance (Miles)'] > 0)]
Total_Visits = Total_Visits_df.groupby('FSM')['Distance (Miles)'].count()
#Average_Distance = df.groupby('FSM')['Distance (Miles)'].mean()
Average_Distance = Total_Visits_df.groupby('FSM')['Distance (Miles)'].mean()
#Average_Time = df.groupby('FSM')['Time (Seconds)'].mean() 
Average_Time = Total_Visits_df.groupby('FSM')['Time (Seconds)'].mean() 
Average_Time = Average_Time / 60
#Average_Time = str(datetime.timedelta(seconds = Average_Time))
fsm_allgrouped = DataFrame(dict(Total_Distance = Total_Distance, Total_Visits = Total_Visits, Total_Time_Minutes = Total_Time, 
	Average_Distance = Average_Distance, Average_Time_Minutes = Average_Time)).reset_index()
fsm_allgrouped.to_csv(summaryoutput, sep=',' )
summaryoutput.close()

#market summary of total time, total distance, and total visits	
market_summaryoutput = open('BingMapsMarketSummary.csv', 'w', newline='')
df = pandas.read_csv('BingMapsOutputDetail.csv', encoding = "ISO-8859-1")
Market_Total_Distance = df.groupby('MARKET')['Distance (Miles)'].sum()
Market_Total_Time = df.groupby('MARKET')['Time (Seconds)'].sum()
Market_Total_Time = Market_Total_Time / 60
#Market_Total_Time = str(datetime.timedelta(seconds = int(Market_Total_Time)))
#Market_Total_Visits = df.groupby('MARKET')['FromAddress'].count()
Market_Total_Visits_df = df[(df['Distance (Miles)'] > 0)]
Market_Total_Visits = Market_Total_Visits_df.groupby('MARKET')['Distance (Miles)'].count()
Market_Average_Distance = Market_Total_Visits_df.groupby('MARKET')['Distance (Miles)'].mean()
Market_Average_Time = Market_Total_Visits_df.groupby('MARKET')['Time (Seconds)'].mean() 
#Market_Average_Distance = df.groupby('MARKET')['Distance (Miles)'].mean()
#Market_Average_Time = df.groupby('MARKET')['Time (Seconds)'].mean()
Market_Average_Time = Market_Average_Time / 60
#Market_Average_Time = str(datetime.timedelta(seconds = int(Market_Average_Time)))
market_allgrouped = DataFrame(dict(Market_Total_Distance = Market_Total_Distance, Market_Total_Visits = Market_Total_Visits, 
	Market_Total_Time_Minutes = Market_Total_Time, Market_Average_Distance = Market_Average_Distance, Market_Average_Time_Minutes = Market_Average_Time)).reset_index()
market_allgrouped.to_csv(market_summaryoutput, sep=',' )
market_summaryoutput.close()	

#region summary of total time, total distance, and total visits	
region_summaryoutput = open('BingMapsRegionSummary.csv', 'w', newline='')
df = pandas.read_csv('BingMapsOutputDetail.csv', encoding = "ISO-8859-1")
Region_Total_Distance = df.groupby('REGION')['Distance (Miles)'].sum()
Region_Total_Time = df.groupby('REGION')['Time (Seconds)'].sum()
Region_Total_Time = Region_Total_Time / 60
#Region_Total_Time = str(datetime.timedelta(seconds = int(Region_Total_Time)))
#Region_Total_Visits = df.groupby('REGION')['FromAddress'].count()
Region_Total_Visits = df[(df['Distance (Miles)'] > 0)]
Region_Total_Visits = Region_Total_Visits.groupby('REGION')['Distance (Miles)'].count()
#Region_Average_Distance = df.groupby('REGION')['Distance (Miles)'].mean()
#Region_Average_Time = df.groupby('REGION')['Time (Seconds)'].mean()
Region_Average_Distance = Total_Visits_df.groupby('REGION')['Distance (Miles)'].mean() 
Region_Average_Time = Total_Visits_df.groupby('REGION')['Time (Seconds)'].mean() 
Region_Average_Time = Region_Average_Time / 60
#Region_Average_Time = str(datetime.timedelta(seconds = int(Region_Average_Time)))
region_allgrouped = DataFrame(dict(Region_Total_Distance = Region_Total_Distance, Region_Total_Visits = Region_Total_Visits, 
	Region_Total_Time_Minutes = Region_Total_Time, Region_Average_Distance = Region_Average_Distance, Region_Average_Time_Minutes = Region_Average_Time)).reset_index()
region_allgrouped.to_csv(region_summaryoutput, sep=',' )
region_summaryoutput.close()	


#Writing summaries to worksheets in xlsx file
region_allgrouped.to_excel(writer, sheet_name='Region Summary')
worksheet = writer.sheets['Region Summary']
market_allgrouped.to_excel(writer, sheet_name='Market Summary')
worksheet = writer.sheets['Market Summary']
fsm_allgrouped.to_excel(writer, sheet_name='FSM Summary')
worksheet = writer.sheets['FSM Summary']
result.to_excel(writer, sheet_name='Details')
worksheet = writer.sheets['Details']


#Deleting csv output files since we have the xlsx file
os.remove('BingMapsOutput.csv')
os.remove('BingMapsOutputDetail.csv')
os.remove('BingMapsFSMSummary.csv')
os.remove('BingMapsRegionSummary.csv')
os.remove('BingMapsMarketSummary.csv')

print("Total processing errors: " + str(errorcount))
print("Complete")