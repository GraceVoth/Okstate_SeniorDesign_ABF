# -*- coding: utf-8 -*-
"""
Created on Fri Sep 23 13:03:38 2022

@author: grace_elizabeth
"""

## INSTRUCTIONS FOR USE ##
# - Specify the name of your input excel file
# - Name the output excel file
# - Select the distance increment in miles between potential service centers - granularity of the grid
# - Input the latitude and longitude values of current service center

## DEBUGGING HELP ##
# - Make sure your excel files follow the template of our excel files
#    (data is pulled in based on sheet names and column headers)
#    (the script is extremely case sensitive)

import pandas as pd
import requests
import time
import os
    
start = time.perf_counter()

# File paths
file_input = "CA_LH_Overall_Data.xlsx"
file_output = "Overall_Output.xlsx"
    
# Distance in miles between grid points (granularity of grid)
g = 5

# Input current service center latitude and longitude coordinates
current_SC_lat = 29.811
current_SC_long = -95.3095

# Import data
absolute_path = os.path.dirname(__file__)
full_path_input = os.path.join(absolute_path, file_input)
df = pd.read_excel(full_path_input, None)
    
time1 = time.perf_counter() - start
print(f"Read Excel: {time1:0.4f}s")
start = time.perf_counter()

# Create grid of possible service center locations
latitude = []                               # all latitude values
longitude = []                              # all longitude values
lat_long = []                               # zipped lat and long coordinates
latitude2 = []                              # unzipped latitude values
longitude2 = []                             # unzipped longitude values
lat_value = min(df["Zips for Grid"].Lat)
max_lat = max(df["Zips for Grid"].Lat)
long_value = min(df["Zips for Grid"].Long)
max_long = max(df["Zips for Grid"].Long)

while lat_value <= (max_lat + g/69):        # latitude values incremented every g miles over service area
    latitude.append(lat_value)
    lat_value += g/69                       # (miles between grid point)/(miles in a degree latitude) https://www.usgs.gov/faqs/how-much-distance-does-degree-minute-and-second-cover-your-maps#:~:text=One%20degree%20of%20latitude%20equals,one%20second%20equals%2080%20feet.

while long_value <= (max_long + g/54.6):    # longitude values incremented every g miles over service area
    longitude.append(long_value)
    long_value += g/54.6                    # (miles between grid point)/(miles in a degree longitude)

for i in range(len(latitude)):
    for j in range(len(longitude)):
        lat_long.append((latitude[i],longitude[j]))
        latitude2.append(latitude[i])
        longitude2.append(longitude[j])
latitude2.append(current_SC_lat)
longitude2.append(current_SC_long)
lat_long.append((current_SC_lat, current_SC_long))

time2 = time.perf_counter() - start
print(f"Create grid: {time2:0.4f}s")
start = time.perf_counter()

# Destination coordinates
dest_lat = df["Overall"]['Lat'].tolist() 
dest_long = df["Overall"]['Long'].tolist()
dest = []
for i in range(len(dest_lat)):
    dest.append((dest_lat[i],dest_long[i])) # combines latitude and longitude in ( , )

time3 = time.perf_counter() - start
print(f"Write lists: {time3:0.4f}s")
start = time.perf_counter()

# Fuction retrieving the route statistics between pickup and dropoff locations
def get_route(pickup_lon, pickup_lat, dropoff_lon, dropoff_lat):
    
    loc = "{},{};{},{}".format(pickup_lon, pickup_lat, dropoff_lon, dropoff_lat)
    url = "http://router.project-osrm.org/route/v1/driving/"
    r = requests.get(url + loc) 
    if r.status_code!= 200:
        return {}
  
    res = r.json()   
    distance = res['routes'][0]['distance']
    duration = res['routes'][0]['duration']
    weight = res['routes'][0]['weight']

    return [distance, duration, weight]

# Create lists of route statitics for all potential service center and destination combinations
d = []
dur = []
w = []
for i in range(len(lat_long)): 
        for j in range(len(dest)): 
            route = get_route(longitude2[i],latitude2[i],dest_long[j],dest_lat[j])
            d.append(route[0])
            dur.append(route[1])
            w.append(route[2])
                
time4 = time.perf_counter() - start
print(f"Calculate distances with OSRM: {time4:0.4f}s")
start = time.perf_counter()

c = len(df["City"]['ShipTot'].tolist())
l = len(df["Linehaul"]['ShipTot'].tolist())
v = len(dest)

d_city = []
dur_city = []
w_city = []
d_linehaul = []
dur_linehaul = []
w_linehaul = []
d_overall = []
dur_overall = []
w_overall = []
for i in range(len(lat_long)):
    d_city.append(d[(i*c + i*l):((i+1)*c + i*l)])
    dur_city.append(dur[(i*c + i*l):((i+1)*c + i*l)])
    w_city.append(w[(i*c + i*l):((i+1)*c + i*l)])
    d_linehaul.append(d[((i+1)*c + i*l):((i+1)*c + (i+1)*l)])
    dur_linehaul.append(dur[((i+1)*c + i*l):((i+1)*c + (i+1)*l)])
    w_linehaul.append(w[((i+1)*c + i*l):((i+1)*c + (i+1)*l)])
    d_overall.append(d[(i*v):((i+1)*v)])
    dur_overall.append(dur[(i*v):((i+1)*v)])
    w_overall.append(w[(i*v):((i+1)*v)])

time5 = time.perf_counter() - start
print(f"Split route statistics: {time5:0.4f}s")
start = time.perf_counter()
 
# Save route statistics to excel file       
df2_d_overall_matrix = pd.DataFrame(d_overall)
df2_dur_overall_matrix = pd.DataFrame(dur_overall)      
df2_w_overall_matrix = pd.DataFrame(w_overall)
df2_d_city_matrix = pd.DataFrame(d_city)
df2_dur_city_matrix = pd.DataFrame(dur_city)
df2_w_city_matrix = pd.DataFrame(w_city)
df2_d_linehaul_matrix = pd.DataFrame(d_linehaul)
df2_dur_linehaul_matrix = pd.DataFrame(dur_linehaul)
df2_w_linehaul_matrix = pd.DataFrame(w_linehaul)
with pd.ExcelWriter(os.path.join(absolute_path, file_output)) as writer:
    df2_d_overall_matrix.to_excel(writer, sheet_name ='Overall Distances')
    df2_dur_overall_matrix.to_excel(writer, sheet_name ='Overall Durations')
    df2_w_overall_matrix.to_excel(writer, sheet_name ='Overall Weights')
    df2_d_city_matrix.to_excel(writer, sheet_name ='City Distances')
    df2_dur_city_matrix.to_excel(writer, sheet_name ='City Durations')
    df2_w_city_matrix.to_excel(writer, sheet_name ='City Weights')
    df2_d_linehaul_matrix.to_excel(writer, sheet_name ='Linehaul Distances')
    df2_dur_linehaul_matrix.to_excel(writer, sheet_name ='Linehaul Durations')
    df2_w_linehaul_matrix.to_excel(writer, sheet_name ='Linehaul Weights')

time6 = time.perf_counter() - start
print(f"Print route statistics: {time6:0.4f}s")