import numpy as np
import pandas as pd
from geopy.geocoders import DataBC
from geopy.extra.rate_limiter import RateLimiter
import geopy
import re
import camp

def ditch_commas(df, col):
    df[col] = df[col].str.replace(r',+', ',', regex=True)
    df[col] = df[col].str.replace(r'(^,.*?)', '', regex=True)
    df[col] = df[col].str.replace(r'( ,)+', '', regex=True)
    df[col] = df[col].str.replace(r'( , )+', '', regex=True)
    df[col] = df[col].str.strip()
    df[col] = df[col].str.rstrip(',')

if __name__ == "__main__":
    # Opens the file (must be deencrypted)
    gsr = pd.read_csv('gsr_assisted_living_feb_25.csv').fillna('')
    hosp = pd.read_csv('hlbc_hospitals.csv').fillna('')
    upcc = pd.read_csv('hlbc_urgentandprimarycarecentres (1).csv').fillna('')
    qfd = pd.read_excel('QFD 2020 - public release - 20210105.xlsx', sheet_name='QFD 2020').fillna('')
    writer = pd.ExcelWriter('flagged_added.xlsx', engine='openpyxl')

    temp = [hosp, upcc, gsr, qfd]

    gsr['BUSINESS_NAME'] = gsr['BUSINESS_NAME'].str.upper()
    hosp['SV_NAME'] = hosp['SV_NAME'].str.upper()
    upcc['SV_NAME'] = upcc['SV_NAME'].str.upper()
    qfd['FACILITY_NAME'] = qfd['FACILITY_NAME'].str.upper()
    camp.separate_columns(gsr, 'BUSINESS_NAME')
    camp.separate_columns(hosp, 'SV_NAME')
    camp.separate_columns(upcc, 'SV_NAME')
    camp.separate_columns(qfd, 'FACILITY_NAME')

# hosp
    address = ['STREET_NUMBER','STREET_NAME','STREET_TYPE','STREET_DIRECTION','CITY','PROVINCE']
    hosp["ADDR_FOR_GEO"] = hosp[address].apply(lambda x: ', '.join(x.dropna()), axis=1)

# UPCC
    upcc["ADDR_FOR_GEO"] = upcc[address].apply(lambda x: ', '.join(x.dropna()), axis=1)

# gsr
    gsr['PROVINCE'] = 'BC'
    address = ['STREET_ADDRESS','CITY','PROVINCE']
    gsr["ADDR_FOR_GEO"] = gsr[address].apply(lambda x: ', '.join(x.dropna()), axis=1)

# QFD
    qfd['PROVINCE'] = 'BC'
    qfd["ADDR_FOR_GEO"] = qfd[address].apply(lambda x: ', '.join(x.dropna()), axis=1)

    # Removes excess commas for geocoding

    for item in temp:
        ditch_commas(item, "ADDR_FOR_GEO")

    # Geocoding
    geolocator = DataBC(user_agent="ISU_filter")
    geocoded = RateLimiter(geolocator.geocode, min_delay_seconds=1/15)


    # Modified
    for item in temp:
        item['GEO_LOCATION'] = item['ADDR_FOR_GEO'].astype(str).apply(lambda doot: geocoded(doot) if doot!= 'nan' else '')
        item['GEO_ADDRESS'] = item['GEO_LOCATION'].apply(lambda loc: loc.address if loc else "")
        item['GEO_RAW'] = item['GEO_LOCATION'].apply(lambda loc: loc.raw if loc else "")
        item['GEO_GPS'] = item['GEO_LOCATION'].apply(lambda loc: loc.point if loc else "")
        item['GEO_LATITUDE'] = item['GEO_GPS'].apply(lambda loc: loc.latitude if loc else "")
        item['GEO_LONGITUDE'] = item['GEO_GPS'].apply(lambda loc: loc.longitude if loc else "")

    # Setting Flags for any generic address given that is within BC and isnt a PO BOX
    for item in temp:
        item['FLAG'] = 0
        item.loc[(item['GEO_ADDRESS'].astype(str).str.contains(r"^[^,]*((,[^,]*){0,1}$)", regex=True, na=True))|~(item['GEO_ADDRESS'].astype(str).str.contains(r"\d", regex=True, na=True)), 'FLAG'] = 1

    # Exports excel file
    hosp.to_excel(writer, sheet_name = 'hosp_geo', index=False)
    upcc.to_excel(writer, sheet_name = 'upcc_geo', index=False)
    gsr.to_excel(writer, sheet_name = 'gsr_geo', index=False)
    qfd.to_excel(writer, sheet_name = 'qfd_geo', index=False)
    writer.save()







    

    
