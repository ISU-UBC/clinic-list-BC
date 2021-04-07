import numpy as np
import pandas as pd
import xlwings as xw
from geopy.geocoders import DataBC
from geopy.extra.rate_limiter import RateLimiter
import geopy
import re


if __name__ == "__main__":
    # Opens the file (must be deencrypted)
    modified = pd.read_excel('flagged_lists_manual_edits.xlsx', sheet_name='Flagged modified')
    corrfac = pd.read_excel('flagged_lists_manual_edits.xlsx', sheet_name='Flagged Corrections Facilities')
    hlbc = pd.read_excel('flagged_lists_manual_edits.xlsx', sheet_name='Flagged HLBC Clinic List')
    writer = pd.ExcelWriter('working_list.xlsx', engine='openpyxl')

    # Looks for Flags in all the lists and isolates them for geocoding 
    modified_partial_1 = modified[(modified['FLAG_1']==1)]
    modified_partial_2 = modified[(modified['FLAG_2']==1)]
    modified_partial_1 = modified_partial_1.dropna(how = 'all')
    modified_partial_2 = modified_partial_2.dropna(how = 'all')
    corrfac_partial = corrfac.where(corrfac['FLAG']==1)
    corrfac_partial = corrfac_partial.dropna(how = 'all')
    hlbc_partial = hlbc.where(hlbc['FLAG']==1)
    hlbc_partial = hlbc_partial.dropna(how = 'all')

    # Geocoding
    geolocator = DataBC(user_agent="ISU_filter")
    geocoded = RateLimiter(geolocator.geocode, min_delay_seconds=1/15)

    # Geocodes HLBC and Corrections Facilities
    lists_for_next_for = [corrfac_partial, hlbc_partial]
    for item in lists_for_next_for:
        item['GEO_LOCATION'] = item['ADDR_FOR_GEO'].astype(str).apply(lambda doot: geocoded(doot) if doot!= 'nan' else '')
        item['GEO_ADDRESS'] = item['GEO_LOCATION'].apply(lambda loc: loc.address if loc else "")
        item['GEO_RAW'] = item['GEO_LOCATION'].apply(lambda loc: loc.raw if loc else "")
        item['GEO_GPS'] = item['GEO_LOCATION'].apply(lambda loc: loc.point if loc else "")
        item['GEO_LATITUDE'] = item['GEO_LOCATION'].apply(lambda loc: loc.latitude if loc else "")
        item['GEO_LONGITUDE'] = item['GEO_LOCATION'].apply(lambda loc: loc.longitude if loc else "")

    # Geocodes Modified
    # FLAG 1
    modified_partial_1['GEO_LOCATION_1'] = modified_partial_1['ADDR_FOR_GEO_1'].astype(str).apply(lambda doot: geocoded(doot) if doot!= 'nan' else '')
    modified_partial_1['GEO_ADDRESS_1'] = modified_partial_1['GEO_LOCATION_1'].apply(lambda loc: loc.address if loc else "")
    modified_partial_1['GEO_RAW_1'] = modified_partial_1['GEO_LOCATION_1'].apply(lambda loc: loc.raw if loc else "")
    modified_partial_1['GEO_GPS_1'] = modified_partial_1['GEO_LOCATION_1'].apply(lambda loc: loc.point if loc else "")
    modified_partial_1['GEO_LATITUDE_1'] = modified_partial_1['GEO_LOCATION_1'].apply(lambda loc: loc.latitude if loc else "")
    modified_partial_1['GEO_LONGITUDE_1'] = modified_partial_1['GEO_LOCATION_1'].apply(lambda loc: loc.longitude if loc else "")
    # FLAG 2
    modified_partial_2['GEO_LOCATION_2'] = modified_partial_2['ADDR_FOR_GEO_2'].astype(str).apply(lambda doot: geocoded(doot) if doot!= 'nan' else '')
    modified_partial_2['GEO_ADDRESS_2'] = modified_partial_2['GEO_LOCATION_2'].apply(lambda loc: loc.address if loc else "")
    modified_partial_2['GEO_RAW_2'] = modified_partial_2['GEO_LOCATION_2'].apply(lambda loc: loc.raw if loc else "")
    modified_partial_2['GEO_GPS_2'] = modified_partial_2['GEO_LOCATION_2'].apply(lambda loc: loc.point if loc else "")
    modified_partial_2['GEO_LATITUDE_2'] = modified_partial_2['GEO_LOCATION_2'].apply(lambda loc: loc.latitude if loc else "")
    modified_partial_2['GEO_LONGITUDE_2'] = modified_partial_2['GEO_LOCATION_2'].apply(lambda loc: loc.longitude if loc else "")

    # Joins final lists and drops duplicates based on identifying factors through:
        # Provider name and MSC
    modified_final = pd.concat([modified, modified_partial_1, modified_partial_2])
    fullname = ['Last Name', 'Given Names', 'Msc']
    modified_final['fullname'] = modified_final[fullname].astype(str).apply(lambda x: ', '.join(x.dropna()), axis=1)
    modified_final = modified_final.drop_duplicates(subset='fullname', keep="last")
    modified_final = modified_final.drop(columns=['fullname'])

        # Corrections Facility Name 
    corrfac_final = pd.concat([corrfac, corrfac_partial])
    corrfac_final = corrfac_final.drop_duplicates(subset='GEO_LOCATION', keep="last")

        # SL_Reference
    hlbc_final = pd.concat([hlbc, hlbc_partial])
    hlbc_final = hlbc_final.drop_duplicates(subset='SL_REFERENCE', keep="last")

    # Exports excel file
    modified_final.to_excel(writer, sheet_name = 'Modified Final', index=False)
    corrfac_final.to_excel(writer, sheet_name = 'Corrections Facilities Final', index=False)
    hlbc_final.to_excel(writer, sheet_name = 'HLBC Final', index=False)
    writer.save()







    

    
