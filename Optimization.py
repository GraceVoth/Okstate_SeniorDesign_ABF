# -*- coding: utf-8 -*-
"""
Created on Fri Sep 23 13:03:38 2022

@author: grace_elizabeth
"""

## INSTRUCTIONS FOR USE ##
# - Select the analysis type (Overall, City, Linehaul)
# - Select the type of cost to be used in the objective function (Distance, Duration, Weight)
# - Input file names (this script and the excel files must all be saved in the same folder)
# - Select the distance increment in miles between potential service centers - granularity of the grid (same as first script)
# - Select the number of service centers you would like to locate in that area
# - Adjust the bounds for the percentage of customers a service center serves if selecting more than one SC

## DEBUGGING HELP ##
# - Make sure your excel files follow the template of our excel files
#    (data is pulled in based on sheet names and column headers)
#    (the script is extremely case sensitive)


from gurobipy import *
import pandas as pd
import requests
import time
import os

try:
    
    # Analysis type
    #name = "Overall"
    #name = "City"
    name = "Linehaul"
    
    # Cost type
    cost = "Distances"                         # road distance
    #cost = "Durations"                        # travel time
    #cost = "Weights" 
 
    # File names
    file_ABF = "CA_LH_Overall_Data.xlsx"
    file_OSRM = "Overall_Output.xlsx"
    
    # Distance in miles between grid points (granularity of grid)
    g = 5
    
    # Number of service centers to be located and bounds for percentage of customers served from each service center
    p = 1
    LB = 1/p - 0.1                             # 0.4 or 40% for p = 2 service centers
    UB = 1/p + 0.1                             # 0.6 or 60% for p = 2 service centers
    
    print(name, "analysis optimizing with travel", cost)
    start = time.perf_counter()
    
    #Import data
    absolute_path = os.path.dirname(__file__)
    full_path_ABF = os.path.join(absolute_path, file_ABF)
    full_path_OSRM = os.path.join(absolute_path, file_OSRM)
    df = pd.read_excel(full_path_ABF, None)
    df2 = pd.read_excel(full_path_OSRM, sheet_name = name + " " + cost, index_col = 0)  # Saved excel of overall analysis distances for quicker optimization
    time1 = time.perf_counter() - start
    print(f"Read Excel sheets: {time1:0.4f}s")
    start = time.perf_counter()
    
    # Create grid of potential service center locations
    latitude = []                               # all latitude values
    longitude = []                              # all longitude values
    lat_long = []                               # zipped lat and long coordinates
    lat_value = min(df["Zips for Grid"].Lat)
    max_lat = max(df["Zips for Grid"].Lat)
    long_value = min(df["Zips for Grid"].Long)
    max_long = max(df["Zips for Grid"].Long)
    
    while lat_value <= (max_lat + g/69):        # latitude values incremented every g miles over service area
        latitude.append(round(lat_value, 4))
        lat_value += g/69                       # (miles between grid point)/(miles in a degree latitude) https://www.usgs.gov/faqs/how-much-distance-does-degree-minute-and-second-cover-your-maps#:~:text=One%20degree%20of%20latitude%20equals,one%20second%20equals%2080%20feet.
    
    while long_value <= (max_long + g/54.6):    # longitude values incremented every g miles over service area
        longitude.append(round(long_value, 4))
        long_value += g/54.6                    # (miles between grid point)/(miles in a degree longitude)
    
    for i in range(len(latitude)):
        for j in range(len(longitude)):
            lat_long.append((latitude[i],longitude[j]))
    lat_long.append((29.811, -95.3095))         # adds the current service center location to the list of potential service centers to be evaluated in the optimization      
    time2 = time.perf_counter() - start
    print(f"Create grid: {time2:0.4f}s")
    start = time.perf_counter()
    
    # Parameters
    h = df[name]['ShipTot'].tolist()
    d = df2
    time5 = time.perf_counter() - start
    print(f"Write lists and arrays: {time5:0.4f}s")
    start = time.perf_counter()
        
    # Indices
    n = len(h) # customer locations i
    s = len(lat_long) # service center locations j
        
    # Write model
    m = Model("Facility Location Optimization")
    
    # Decision variables
    x = m.addVars(range(s), vtype = GRB.BINARY, name = "Service Center Selected")
    y = m.addVars(range(n), range(s), vtype = GRB.BINARY, name = "Assigment")
    pa = m.addVars(range(s), lb = 0, vtype = GRB.CONTINUOUS, name = "Percent Assigned")
    
    # Objective function
    m.setObjective(quicksum(h[i] * d[i][j] * y[i,j] for i in range(n) for j in range(s)), GRB.MINIMIZE)
    
    # Constraints
    m.addConstr(quicksum(x[j] for j in range(s)) == p, name = "1 Center Located") 
    for i in range(n):
        m.addConstr(quicksum(y[i,j] for j in range(s)) == 1, name = "1 Customer - 1 Service Center")
        for j in range(s):
            m.addConstr(y[i,j] - x[j] <= 0, name = "Assignment to Selected Service Centers")        
    
    for j in range(s):
        m.addConstr(quicksum(y[i,j] for i in range(n))/n >= LB*x[j], name = "Percent of customers assigned to SC greater than lower bound")
        m.addConstr(quicksum(y[i,j] for i in range(n))/(-n) >= -UB*x[j], name = "Percent of customers assigned to SC less than upper bound")
        
    # Gurobi Optimizer
    m.optimize()
    if m.status == GRB.OPTIMAL:
        time8 = time.perf_counter() - start
        print(f"Optimize Time: {time8:0.4f}s")
        for j in range(s):
                if x[j].x > 0:
                    percent_assigned = sum(y[i,j].x for i in range(n))/n
                    print("Service Center Selected: #", j, lat_long[j], round(percent_assigned*100), "% assigned")                    
        if cost == "Distances":
            print(f'Objective = {m.objVal:0.2f} meters')
        if cost == "Durations":
            print(f'Objective = {m.objVal:0.2f} seconds')
        if cost == "Weights":
            print(f'Objective = {m.objVal:0.2f} units?')
    elif m.status == GRB.INFEASIBLE:
       print('LP is infeasible.')
    elif m.status == GRB.UNBOUNDED:
       print('LP is unbounded.')
except GurobiError:
    print('Error reported') 