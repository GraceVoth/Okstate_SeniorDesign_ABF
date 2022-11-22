# Okstate_SeniorDesign_ABF_F22

This repository holds the python scripts and excel template input file used for the facility location optimization project done for ABF Freight Fall 2022 by Oklahoma State's IEM Senior Design Team: Caleb Triplett, Darcie Golden, and Grace Voth.

The first python script, OSRM_Integration, connects to OSRM and writes all the distance calculations to an excel sheet to be used in the optimization python script. This script takes approximately 40 hours to run with Houston service area data.
  To use OSRM_Integration:
  1.	Specify the name of your input excel file
  2.	Name the output excel file
  3.	Select the distance increment in miles between potential SCs - granularity of the grid
  4.	Input the latitude and longitude values of current SC

The second python script, Optimization, finds the optimal SC location using Gurobi. This script takes anywhere from 30 seconds to 5 minutes to run with Houston service area data.
  To use Optimization:
  1.	Select the analysis type (Overall, City, Linehaul)
  2.	Select the type of cost to be used in the objective function (Distance, Duration, Weight)
  3.	Input file names (this script and the excel files must all be saved in the same folder)
  4.	Select the distance increment in miles between potential SCs - granularity of the grid (same as first script)
  5.	Select the number of SCs you would like to locate in that area
  6.	Adjust the bounds for the percentage of customers a SC serves if selecting more than one SC
*Note: Both these scripts rely heavily on the structure of the input excel files, data is called into the script using specific sheet names and column headers. Be sure that your excel input file follows the structure of our template file.*
