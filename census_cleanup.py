# import libraries 

import geopandas as gp
import pandas as pd
from shapely.geometry import Point

# load files 
NTA_data = gp.read_file("nynta2020.shp")
Census_data = pd.read_excel("nyc_decennialcensusdata_2010_2020_change.xlsx",sheet_name='2010, 2020, and Change', index_col ='Geography')

# view entire table
pd.set_option('display.max_columns', 50) #replace n with the number of columns you want to see completely
pd.set_option('display.max_rows', 167)

#clean up empty columns
Census_cleanup = Census_data.drop(Census_data.index[1])

del Census_cleanup['Unnamed: 2']
Census_cleanup2 = Census_cleanup.drop(Census_cleanup.columns[[2]], axis=1)

# Cenus data by GeoType : NTA2020 
C_DF = Census_cleanup2[Census_cleanup2['Unnamed: 1'].isin(['NTA2020'])]

# Cenus data by GeoType : NTA2020 : name (sorted)
C_Name = C_DF.sort_values(by=['Unnamed: 5'])
C_Name


# isolates just GeoType and Name
Census_clean = C_Name.loc[:, C_Name.columns.intersection(['Unnamed: 1','Unnamed: 5'])]
Census_clean

# rename unnamed columns to GeoType and Name
Census_clean.rename(columns = {'Unnamed: 1':'GeoType', 'Unnamed: 5':'Name'}, inplace = True)
Census_clean

Cleaned_Census = Census_clean.to_csv(r'C:\Users\mvento\Desktop\Python scripts\cleaned_census.csv')
Cleaned_Census2 = pd.read_csv(r'C:\Users\mvento\Desktop\Python scripts\cleaned_census.csv')

print("CSV completed ...")

