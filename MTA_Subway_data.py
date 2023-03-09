# Melissa Vento
# MTA_Subway_data.py
# Objective : create a python script that can automatically download the data, do the appropriate joins, and output the results into a csv file
# Power BI document: used to visualize for correctness
# Z:\PMA\Maria_Rodriguez\Asset Management Database\Shapefiles\MTA\google_transit\testing_mta_subway dashboard.pbix

# Step 1 : Automatically download data using a website
#   file url : http://web.mta.info/developers/data/nyct/subway/google_transit.zip
# Step 2 : Do appropriate joins 
#    I did here is the MS Access database where I was testing out my queries: 
#       Z:\PMA\Maria_Rodriguez\Asset Management Database\Shapefiles\MTA\google_transit\testing_mta_subway.accdb
#    Power BI document:
#       Z:\PMA\Maria_Rodriguez\Asset Management Database\Shapefiles\MTA\google_transit\testing_mta_subway dashboard.pbix
# Step 3 : Output results into a csv file


# import libraries
#  Pandas   : used for data structures and operations for manipulating numerical tables and time series.
#  requests : to download file from internet (URL)
#  zipfile  : used to upzip the file we downloaded
import pandas as pd 
import requests 
from zipfile import ZipFile

# Step 1 : Automatically download data using a website
download_url = r'http://web.mta.info/developers/data/nyct/subway/google_transit.zip'

req = requests.get(download_url)
filename = req.url[download_url.rfind('/')+1:]

with open(filename, 'wb') as f:
    for chunk in req.iter_content(chunk_size = 8192):
        if chunk:
            f.write(chunk)

def download_file(url, filename = ''):
    try:
        if filename:
            pass
        else:
            filename = req.url[download_url.rfind('/')+1:]
        
        with requests.get(url) as req:
            with open(filename, 'wb') as f:
                for chunk in req.iter_content(chunk_size= 8192):
                    if chunk:
                        f.write(chunk)
            return filename
    except Exception as e:
        print(e)
        return None

download_file(download_url, 'google_transit_py.zip')
print("Download from URL successful")                     

# Step 1A : unzip downloaded file
# unzip downloaded file 
with ZipFile('google_transit_py.zip', 'r') as zipObj:
   # Extract all the contents of zip file in current directory
   zipObj.extractall()
print("Unzip downloaded file completed...")                     


# Step 1 : Complete



# Step 2 : Do appropriate joins 

# I did here is the MS Access database where I was testing out my queries: 
# Z:\PMA\Maria_Rodriguez\Asset Management Database\Shapefiles\MTA\google_transit\testing_mta_subway.accdb

# 1st Join
# name : qryTripsStopTimes
# SELECT DISTINCT Stop_times.trip_id, Stop_times.stop_id, Trips.route_id, Trips.trip_headsign, Trips.shape_id
# FROM Stop_times INNER JOIN Trips ON Stop_times.trip_id = Trips.trip_id;
# INNER JOIN Trips 

    # read Stop_times.txt file 
Stop_times = pd.read_csv("Stop_times.txt")

# read Trips.txt file 
Trips = pd.read_csv("Trips.txt")

#INNER JOIN Stop_times & Trips
#  Using Stop_times
# Stop_times.trip_id = Trips.trip_id;

# possibly a better way to do this.  For loop? inner merge?  need more time to optimize this code

# merge both tables
stop_trip_time = pd.merge(Stop_times, Trips, on = "trip_id")

# drop following columns to only get routes, stop_times information
stop_trip_time_ = stop_trip_time.drop(columns = ["arrival_time", "departure_time", "stop_sequence", "stop_headsign", "pickup_type", "drop_off_type", "shape_dist_traveled", "service_id", "direction_id", "block_id" ])

# write merged table to csv
stop_trip_time_.to_csv('qryTripsStopTimes.csv',index=False)
print('Successfully created qryTripsStopTimes CSV : 1st JOIN Complete')

# 1st Join Complete

# 2nd Join
# name: qryDistinctTripStopTimes
# SELECT DISTINCT qryTripsStopTimes.stop_id, qryTripsStopTimes.route_id
# FROM qryTripsStopTimes;
#getting records qryTripsStopTimes

# read qryTripsStopTimes results from above 
qryTripsStopTimes = pd.read_csv("qryTripsStopTimes.csv")

# extract the following 2 columns : "stop_id", "route_id"
stop_route = qryTripsStopTimes[["stop_id", "route_id"]]


#write stop_route new table to csv  
stop_route.to_csv('qryDistinctTripStopTimes.csv',index=False)
print('Successfully created qryDistinctTripStopTimes CSV : 2nd JOIN Complete')

# 2nd Join Complete


#3rd Join
# name: qryDistinctTripStopTimes_Routes
#SELECT DISTINCT qryDistinctTripStopTimes.stop_id, qryDistinctTripStopTimes.route_id, Routes.route_long_name
# FROM qryDistinctTripStopTimes INNER JOIN Routes ON qryDistinctTripStopTimes.route_id = Routes.route_id;
# INNER JOIN Routes
# using qryDistinctTripStopTimes

# read qryDistinctTripStopTimes results from above 
qryDistinctTripStopTimes = pd.read_csv("qryDistinctTripStopTimes.csv")

# read Routes.txt 
Routes = pd.read_csv("Routes.txt")

# merge  using INNER JOIN
Stop_time_Routes = pd.merge(qryDistinctTripStopTimes, Routes, on = ['route_id'], how = 'inner') 

# drop columns that are not needed
Stop_time_Routes_ = Stop_time_Routes.drop(columns = ["agency_id", "route_short_name", "route_desc", "route_type", "route_url", "route_color", "route_text_color" ])



#write Stop_time_Routes new table to csv  
Stop_time_Routes_.to_csv('qryDistinctTripStopTimes_Routes.csv',index=False)
print('Successfully created qryDistinctTripStopTimes_Routes CSV : 3rd JOIN Complete')

# 3rd Join Complete


# 4th final Join
# name : final
#SELECT DISTINCT Stops.parent_station, Stops.stop_name, Stops.stop_lat, Stops.stop_lon, qryDistinctTripStopTimes_Routes.route_id, qryDistinctTripStopTimes_Routes.route_long_name
#FROM qryDistinctTripStopTimes_Routes INNER JOIN Stops ON qryDistinctTripStopTimes_Routes.stop_id = Stops.stop_id;
# INNER JOIN Stops
# using qryDistinctTripStopTimes_Routes

# read qryDistinctTripStopTimes results from above 
qryDistinctTripStopTimes_Routes = pd.read_csv("qryDistinctTripStopTimes_Routes.csv")

Stops = pd.read_csv("Stops.txt")

# merge using INNER JOIN
Stops_Routes = pd.merge(Stops, qryDistinctTripStopTimes_Routes, on = ['stop_id'], how = 'inner') 
Stops_Routes

# drop columns that are not needed
Stops_Routes_ = Stops_Routes.drop(columns = ["location_type", "stop_url", "zone_id", "stop_desc", "stop_code", "stop_id"])

# shift column 'parent_station' to first position
first_column = Stops_Routes_.pop('parent_station')

# insert column using insert(position,column_name,first_column) function
Stops_Routes_.insert(0, 'parent_station', first_column)


# Find unique values of a column
final = Stops_Routes_.drop_duplicates()

# Step 3 : Output results into a csv file

#write Stop_time_Routes new table to csv  
final.to_csv('final.csv',index=False)
print('Successfully created final CSV : 4th & final JOIN Complete')
print('Python Script complete')
