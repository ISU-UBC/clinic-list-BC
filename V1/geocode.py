import numpy as np
import pandas as pd
from geopy.geocoders import DataBC
from geopy.extra.rate_limiter import RateLimiter
import geopy
import re

def ditch_commas(df, col):
    df[col] = df[col].str.replace(r',+', ',', regex=True)
    df[col] = df[col].str.replace(r'(^,.*?)', '', regex=True)
    df[col] = df[col].str.replace(r'( ,)+', '', regex=True)
    df[col] = df[col].str.replace(r'( , )+', '', regex=True)
    df[col] = df[col].str.strip()
    df[col] = df[col].str.rstrip(',')

if __name__ == "__main__":
    # Opens the file (must be deencrypted)
    full_list = pd.read_excel('modified.xlsx', sheet_name='modified')
    hlbc_list = pd.read_excel('modified.xlsx', sheet_name='HLBC Clinic List')
    corr_fac_list = pd.read_excel('modified.xlsx', sheet_name='Corrections Facilities')
    writer = pd.ExcelWriter('flagged_lists.xlsx', engine='openpyxl')

    # Joins columns for geocoding
    address = ['STREET_NUMBER','CITY','PROVINCE']
    corr_fac_list["ADDR_FOR_GEO"] = corr_fac_list[address].apply(lambda x: ', '.join(x.dropna()), axis=1)
    address = ['STREET_NUMBER','STREET_NAME','STREET_TYPE','STREET_DIRECTION','CITY','PROVINCE']
    hlbc_list["ADDR_FOR_GEO"] = hlbc_list[address].apply(lambda x: ', '.join(x.dropna()), axis=1)

    # Removes excess commas for geocoding
    ditch_commas(corr_fac_list, "ADDR_FOR_GEO")
    ditch_commas(hlbc_list, "ADDR_FOR_GEO")

    # Geocoding
    geolocator = DataBC(user_agent="ISU_filter")
    geocoded = RateLimiter(geolocator.geocode, min_delay_seconds=1/15)

    # # Test Case
    # test = pd.DataFrame(["ADDR_FOR_GEO_1"])
    # # reg address
    # # test["ADDR_FOR_GEO_1"] = full_list["ADDR_FOR_GEO_1"].iloc[1]
    # # empty address
    # test["ADDR_FOR_GEO_1"] = full_list["ADDR_FOR_GEO_2"].iloc[0]
    # print(test)
    # test['location'] = test["ADDR_FOR_GEO_1"].astype(str).apply(lambda doot: geocoded(doot) if doot!= 'nan' else '')
    # # test['location'] = test["ADDR_FOR_GEO_1"].apply(lambda doot: print(doot, type(doot)) if doot!= 'nan' else '')
    # print(test['location'])
    # test['address_geo'] = test['location'].apply(lambda loc: loc.address if loc else "")
    # test['raw'] = test['location'].apply(lambda loc: loc.raw if loc else "")
    # test['point'] = test['location'].apply(lambda loc: loc.point if loc else "")
    # test['latitude'] = test['point'].apply(lambda loc: loc.latitude if loc else "")
    # test['longitude'] = test['point'].apply(lambda loc: loc.longitude if loc else "")
    # print(test)
    # test['FLAG'] = 0
    # test.loc[test["ADDR_FOR_GEO_1"].astype(str).str.contains(r"^[^,]*((,[^,]*){0,1}$)", regex=True, na=True), 'FLAG'] = 1
    # print(test)
    # test.to_excel('test.xlsx')

    # Modified
    print("Geocoding CPSBC List")
    for j in range(1, 3):
        full_list[f'GEO_LOCATION_{j}'] = full_list[f'ADDR_FOR_GEO_{j}'].astype(str).apply(lambda doot: geocoded(doot) if doot!= 'nan' else '')
        full_list[f'GEO_ADDRESS_{j}'] = full_list[f'GEO_LOCATION_{j}'].apply(lambda loc: loc.address if loc else "")
        full_list[f'GEO_RAW_{j}'] = full_list[f'GEO_LOCATION_{j}'].apply(lambda loc: loc.raw if loc else "")
        full_list[f'GEO_GPS_{j}'] = full_list[f'GEO_LOCATION_{j}'].apply(lambda loc: loc.point if loc else "")
        full_list[f'GEO_LATITUDE_{j}'] = full_list[f'GEO_GPS_{j}'].apply(lambda loc: loc.latitude if loc else "")
        full_list[f'GEO_LONGITUDE_{j}'] = full_list[f'GEO_GPS_{j}'].apply(lambda loc: loc.longitude if loc else "")

    # Corr_facs
    print("Geocoding Corrections Facilities List")
    corr_fac_list['GEO_LOCATION'] = corr_fac_list['ADDR_FOR_GEO'].astype(str).apply(lambda doot: geocoded(doot) if doot!= 'nan' else '')
    corr_fac_list['GEO_ADDRESS'] = corr_fac_list['GEO_LOCATION'].apply(lambda loc: loc.address if loc else "")
    corr_fac_list['GEO_RAW'] = corr_fac_list['GEO_LOCATION'].apply(lambda loc: loc.raw if loc else "")
    corr_fac_list['GEO_GPS'] = corr_fac_list['GEO_LOCATION'].apply(lambda loc: loc.point if loc else "")
    corr_fac_list['GEO_LATITUDE'] = corr_fac_list['GEO_GPS'].apply(lambda loc: loc.latitude if loc else "")
    corr_fac_list['GEO_LONGITUDE'] = corr_fac_list['GEO_GPS'].apply(lambda loc: loc.longitude if loc else "")

    # HLBC 
    print("Geocoding HLBC List")
    hlbc_list['GEO_LOCATION'] = hlbc_list['ADDR_FOR_GEO'].astype(str).apply(lambda doot: geocoded(doot) if doot!= 'nan' else '')
    hlbc_list['GEO_ADDRESS'] = hlbc_list['GEO_LOCATION'].apply(lambda loc: loc.address if loc else "")
    hlbc_list['GEO_RAW'] = hlbc_list['GEO_LOCATION'].apply(lambda loc: loc.raw if loc else "")
    hlbc_list['GEO_GPS'] = hlbc_list['GEO_LOCATION'].apply(lambda loc: loc.point if loc else "")
    hlbc_list['GEO_LATITUDE'] = hlbc_list['GEO_GPS'].apply(lambda loc: loc.latitude if loc else "")
    hlbc_list['GEO_LONGITUDE'] = hlbc_list['GEO_GPS'].apply(lambda loc: loc.longitude if loc else "")

    # Setting Flags for any generic address given that is within BC and isnt a PO BOX
    corr_fac_list['FLAG'] = 0
    hlbc_list['FLAG'] = 0
    corr_fac_list.loc[corr_fac_list['GEO_ADDRESS'].astype(str).str.contains(r"^[^,]*((,[^,]*){0,1}$)", regex=True, na=True), 'FLAG'] = 1
    hlbc_list.loc[hlbc_list['GEO_ADDRESS'].astype(str).str.contains(r"^[^,]*((,[^,]*){0,1}$)", regex=True, na=True), 'FLAG'] = 1
    for i in range(1,3):
        full_list[f'FLAG_{i}'] = 0
        full_list.loc[(full_list[f'GEO_ADDRESS_{i}'].astype(str).str.contains(r"^[^,]*((,[^,]*){0,1}$)", regex=True, na=True))|~(full_list[f'GEO_ADDRESS_{i}'].astype(str).str.contains(r"\d", regex=True, na=True)), f'FLAG_{i}'] = 1
        full_list.loc[(full_list[f'PO_BOX_ADDR {i}'].str.contains(r".+", regex=True, na=False))&(full_list[f'FLAG_{i}']==1), [f'GEO_LOCATION_{i}',f'GEO_ADDRESS_{i}',f'GEO_RAW_{i}',f'GEO_GPS_{i}',f'GEO_LATITUDE_{i}',f'GEO_LONGITUDE_{i}']] = np.nan
        full_list.loc[(full_list[f'PO_BOX_ADDR {i}'].str.contains(r".+", regex=True, na=False))&(full_list[f'FLAG_{i}']==1), f'FLAG_{i}'] = 0
        full_list.loc[(full_list[f'Province {i}']!='BC')|~(full_list['Specialties & Certifications'].str.contains(r'^Fam.*', regex=True, na=False)), f'FLAG_{i}'] = 0

    # Exports excel file
    full_list.to_excel(writer, sheet_name = 'Flagged modified', index=False)
    hlbc_list.to_excel(writer, sheet_name = 'Flagged HLBC Clinic List', index=False)
    corr_fac_list.to_excel(writer, sheet_name = 'Flagged Corrections Facilities', index=False)
    writer.save()







    

    
