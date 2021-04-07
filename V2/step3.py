# Placeholder for credits

import numpy as np
import pandas as pd
from geopy.geocoders import DataBC
from geopy.extra.rate_limiter import RateLimiter
import geopy
import re
import step1 as s1

class Modified(s1.List):

    def __init__(self, filename, sheetname=None):
        s1.List.__init__(self, filename=filename, sheetname=sheetname)

    def dict_application(self, key, value):
        def create_dict(newdict, key, value):
            newdict.update(pd.Series(self.list.loc[(self.list[key].notna())&(self.list[value].notna()), key].values,index=self.list.loc[(self.list[key].notna())&(self.list[value].notna()), value]).to_dict())
            newdict = {k: v for k, v in newdict.items() if pd.Series(v).notna().all()}  
            return newdict
        def apply_dict(newdict, key, value):
            CPSBC.list = CPSBC.list.reset_index(drop=True)
            CPSBC.list[key] = CPSBC.list[value].map(newdict).fillna(CPSBC.list[key])
        
        newdict={}
        newdict = create_dict(newdict, key, value)
        apply_dict(newdict, key, value)
       
    def geocode_flags(self, geolocator, geocoded, CPSBC_mod=False):
        partial = self.list[(self.list['FLAG']==1)]
        partial.dropna(how = 'all', inplace=True)
        partial['GEO_LOCATION'] = partial['ADDR_FOR_GEO'].astype(str).apply(lambda doot: geocoded(doot) if doot!= 'nan' else '')
        partial['GEO_ADDRESS'] = partial['GEO_LOCATION'].apply(lambda loc: loc.address if loc else '')
        partial['GEO_RAW'] = partial['GEO_LOCATION'].apply(lambda loc: loc.raw if loc else '')
        partial['GEO_GPS'] = partial['GEO_LOCATION'].apply(lambda loc: loc.point if loc else '')
        partial['GEO_LATITUDE'] = partial['GEO_GPS'].apply(lambda loc: loc.latitude if loc else '')
        partial['GEO_LONGITUDE'] = partial['GEO_GPS'].apply(lambda loc: loc.longitude if loc else '')
        self.list = pd.concat([self.list, partial])
        if CPSBC_mod == True:
            fullname = ['Last Name', 'Given Names', 'Postal']
            self.join_columns('fullname', fullname)
            self.list = self.list.drop_duplicates(subset='fullname', keep='last')
            self.list = self.list.drop(columns=['fullname'])
        else:
            self.list = self.list.drop_duplicates(subset='GEO_LOCATION', keep='last')

    def remove_units(self):
        # Removes unit numbers from addresses
        self.list['GEO_SUITE'] = self.list['GEO_ADDRESS'].str.extract(r'(.*)\b\s--\s', expand=False)
        self.list['GEO_ADDRESS'] = self.list['GEO_ADDRESS'].str.replace(r'(.*)(\b\s--\s)', '', regex=True)
        self.list['GEO_LOCATION'] = self.list['GEO_LOCATION'].str.replace(r'(^|.*?)(\D*)(\s--\s)', r'\1\3', regex=True)
        self.list['GEO_LOCATION'] = self.list['GEO_LOCATION'].str.replace(r'(^|\D*?\b\s)(\w*?)(\s--\s)', r'\2\3', regex=True)

    def po_fix(self):
    # PO Boxes prove a unique challenge for geocoding - this provides a unique method of addressing it
        address = ['PO_BOX', 'CITY', 'PROVINCE', 'POSTAL']
        self.list.loc[~(self.list['GEO_ADDRESS'].str.contains(r'(\d+)',regex=True, na=False))&(self.list['PROVINCE']=='BC'), ['GEO_ADDRESS','GEO_LOCATION']] = self.list[address].apply(lambda x: ', '.join(x.dropna()), axis=1) 
        self.remove_commas('GEO_ADDRESS')
        self.remove_commas('GEO_LOCATION')

    def find_duplicate_add(self):
        self.list['temp'] = CPSBC.list['LAST NAME'] + ', ' + CPSBC.list['GIVEN NAMES']
        self.list.loc[self.list.duplicated(subset=['temp', 'GEO_LOCATION'], keep=False), 'ADD_LISTING'] = '1&2'

    def misc_addons(self, categories, renames={}, empties=[], CPSBC_mod=False):
        self.list.rename(columns=renames, inplace=True)
        for item in empties:
            self.empty_col(item)
        
        if CPSBC_mod==False:
            temp = [item for item in categories]
            temp.append('STREET_NUMBER')
            self.list['ADDRESS_FULL'] = self.list[temp].apply(lambda x: ', '.join(x.dropna()), axis=1)
            self.remove_commas(col = 'ADDRESS_FULL')
            self.empty_col('BRACKETS')
            self.list['COUNTRY'] = 'CANADA'

        self.empty_col('CHSA_NUM', 'CHSA_NAME', 'HA_ID', 'HA_NAME', 'CHSA_UR_CLASS')
        self.zero_col('IS_WALKIN', 'IS_UPCC')
        self.remove_units()

    def list_application(self, categories, list_type=None):
        final_headings = [
                'PO_BOX',
                'ADDRESS_FULL',
                'CITY',
                'PROVINCE',
                'COUNTRY',                
                'POSTAL',
                'PHONE', 
                'FAX',
                'OTHER', 
                'CORRECTIONS',
                'FIRST_NATIONS',
                'ADMIN',                
                'HOSP',
                'MHSU',
                'LTC',
                'DERM', 
                'REHAB', 
                'PLASTICS', 
                'PALLIATIVE', 
                'GYN', 
                'PAIN', 
                'GERIATRIC', 
                'IMAGE', 
                'SEX', 
                'ONCO', 
                'DISORDER', 
                'CNS', 
                'SURG', 
                'ORTHO',  
                'TRAVEL', 
                'VASCULAR',
                'VIRTUAL',
                'AGENCY',
                'UNI',
                'CLINIC',
                'CENTRE',
                'MED',
                'SOCIETY',
                'UNIT',
                'HEALTH',
                'FAMILY',
                'BRACKETS',
                'WEBSITE',
                'GEO_SUITE',
                'GEO_LOCATION',
                'GEO_ADDRESS', 
                'GEO_LATITUDE', 
                'GEO_LONGITUDE', 
                'GEO_GPS',
                'CHSA_NUM',
                'CHSA_NAME',
                'HA_ID',
                'HA_NAME',
                'CHSA_UR_CLASS',
                'PHYSICIAN_NAMES',
                'NUM_PHYSICIANS',
                'IS_UPCC',
                'IS_WALKIN'
                ]

        def heading_standardization():
            self.list.columns = self.list.columns.str.upper()
            self.list = self.list[final_headings]

        def categorization():
            if list_type == None:
                for item in categories:
                    temp = 'IS_' + item
                    self.list[temp] = 0
                    self.list.loc[self.list[item].str.contains(r"([0-9A-Z])",regex=True, na=False), temp] = 1
            else:
                formatted_categories = ['IS_' + item for item in categories]
                if (list_type == 'UPCC') | (list_type == 'WALKIN'):
                    temp = 'IS_'+list_type
                    self.list[temp] = 1
                for item in formatted_categories:
                    if item == 'IS_'+list_type:
                        self.list[item] = 1
                    else:
                        self.list[item] = 0

        heading_standardization()
        categorization()
        if list_type != None:
            CPSBC.list = pd.concat([CPSBC.list, self.list]).dropna(how='all')

def assign_id(categories):
    def id_generation(working_list, boolean, droplist, id_name, sheet_name):
        temp = working_list.loc[boolean]
        working_list = working_list.loc[~boolean]
        temp.reset_index(drop=True, inplace=True)
        working_list.reset_index(drop=True, inplace=True)
        temp['ID'] = 1 + temp.index
        temp['ID'] = temp['ID'].apply('{:0>4}'.format)
        temp['ID'] = id_name + temp['ID'].astype(str)
        temp2 = temp.drop(columns = droplist)
        temp2.to_excel(writer, sheet_name = sheet_name, index=False)
        CPSBC.list = pd.concat([CPSBC.list, temp], axis=0)
        return working_list

    droplist = ['IS_' + item for item in categories]
    droplist.append('IS_WALKIN')
    droplist.append('IS_UPCC')
    CPSBC.empty_col('ID')
    working_list = CPSBC.list
    working_list = id_generation(working_list, boolean=(working_list['IS_HOSP']==1), droplist = droplist, id_name='HOS_', sheet_name='Hospital ID') 
    working_list = id_generation(working_list, boolean=(working_list['IS_WALKIN']==1), droplist = droplist, id_name='WIC_', sheet_name='Walk-in Clinic ID') 
    working_list = id_generation(working_list, boolean=(working_list['IS_LTC']==1), droplist = droplist, id_name='LTC_', sheet_name='Long Term Care ID')
    working_list = id_generation(working_list, boolean=(working_list['IS_UPCC']==1), droplist = droplist, id_name='UPC_', sheet_name='Urgent Primary Care Clinic ID')
    working_list = id_generation(working_list, boolean=(working_list['IS_FAMILY']==1), droplist = droplist, id_name='FAM_', sheet_name='Family ID') 
    working_list = id_generation(working_list, boolean=((working_list['IS_CORRECTIONS']==1)), droplist = droplist, id_name='CFI_', sheet_name='Corrections ID')
    working_list = id_generation(working_list, boolean=(working_list['IS_FIRST_NATIONS']==1), droplist = droplist, id_name='FN_', sheet_name='First Nations ID')
    working_list = id_generation(working_list, boolean=(working_list['IS_SEX']==1), droplist = droplist, id_name='SXH_', sheet_name='Sexual Health Clinic ID') 
    working_list = id_generation(working_list, boolean=(working_list['IS_GYN']==1), droplist = droplist, id_name='GYN_', sheet_name='Womens Health Clinic ID')
    working_list = id_generation(working_list, boolean=(working_list['IS_VIRTUAL']==1), droplist = droplist, id_name='VIR_', sheet_name='Virtual ID') 
    working_list = id_generation(working_list, boolean=(working_list['IS_ADMIN']==1), droplist = droplist, id_name='ADM_', sheet_name='Administrative ID')
    specialty_bool = (working_list['IS_MHSU']==1)|(working_list['IS_DERM']==1)|(working_list['IS_REHAB']==1)|(working_list['IS_PLASTICS']==1)|(working_list['IS_PALLIATIVE']==1)|(working_list['IS_PAIN']==1)|(working_list['IS_GERIATRIC']==1)|(working_list['IS_IMAGE']==1)|(working_list['IS_ONCO']==1)|(working_list['IS_DISORDER']==1)|(working_list['IS_CNS']==1)|(working_list['IS_SURG']==1)|(working_list['IS_ORTHO']==1)|(working_list['IS_VASCULAR']==1)|(working_list['IS_TRAVEL']==1)
    working_list = id_generation(working_list, boolean=(specialty_bool), droplist = droplist, id_name='SPC_', sheet_name='Specialty Clinic ID') 
    working_list = id_generation(working_list, boolean=((working_list['IS_CLINIC']==1)|(working_list['IS_CENTRE']==1)), droplist = droplist, id_name='PCC_', sheet_name='Clinic or Centre Not Walk-in ID')
    working_list = id_generation(working_list, boolean=(working_list['NUM_PHYSICIANS']>1), droplist = droplist, id_name='UNM_', sheet_name='Unknown Multi Prac ID')
    working_list['ID'] = 1 + working_list.index
    working_list['ID'] = working_list['ID'].apply('{:0>4}'.format)
    working_list['ID'] = 'UNS_' + working_list['ID'].astype(str)
    working_list.drop(columns = droplist, inplace = True)
    working_list.to_excel(writer, sheet_name = 'Unknown Single Prac ID', index=False)
    CPSBC.list = pd.concat([CPSBC.list, working_list], axis=0)
    CPSBC.list.drop(columns = droplist, inplace = True)
    CPSBC.list = CPSBC.list[CPSBC.list['ID']!='']

def upper_it_all():
    CPSBC.list.columns = CPSBC.list.columns.str.upper()
    WALKIN.list.columns = WALKIN.list.columns.str.upper()
    CORR_FAC.list.columns = CORR_FAC.list.columns.str.upper()
    GSR.list.columns = GSR.list.columns.str.upper()
    HOSP.list.columns = HOSP.list.columns.str.upper()
    UPCC.list.columns = UPCC.list.columns.str.upper()
    QFD.list.columns = QFD.list.columns.str.upper()
    CHSA.list.columns = CHSA.list.columns.str.upper()
    REGION.list.columns = REGION.list.columns.str.upper() 

def apply_physicians():
    CPSBC.list['PHYSICIAN_NAMES'] = CPSBC.list['LAST NAME'] + ', ' + CPSBC.list['GIVEN NAMES'] + ' - ' + CPSBC.list['ADD_LISTING'].astype(str)
    doc_dict_new = {}
    doc_dict = {}
    doc_dict.update(pd.Series(CPSBC.list[f'GEO_LOCATION'].values,index=CPSBC.list['PHYSICIAN_NAMES']).to_dict())
    # keys = docs, values = geo_location
    for key, value in doc_dict.items():
        doc_dict_new[value] = []
    for key, value in doc_dict.items():
        doc_dict_new[value].extend([key]) 
    CPSBC.list['PHYSICIAN_NAMES'] = CPSBC.list['GEO_LOCATION'].map(doc_dict_new).fillna(CPSBC.list['PHYSICIAN_NAMES'])
    CPSBC.list['NUM_PHYSICIANS'] = CPSBC.list['PHYSICIAN_NAMES'].str.len()
    WALKIN.list['PHYSICIAN_NAMES'] = WALKIN.list['GEO_LOCATION'].map(doc_dict_new).fillna('')
    WALKIN.list['NUM_PHYSICIANS'] = WALKIN.list['PHYSICIAN_NAMES'].str.len()
    CORR_FAC.list['PHYSICIAN_NAMES'] = CORR_FAC.list['GEO_LOCATION'].map(doc_dict_new).fillna('')
    CORR_FAC.list['NUM_PHYSICIANS'] = CORR_FAC.list['PHYSICIAN_NAMES'].str.len()
    GSR.list['PHYSICIAN_NAMES'] = GSR.list['GEO_LOCATION'].map(doc_dict_new).fillna('')
    GSR.list['NUM_PHYSICIANS'] = GSR.list['PHYSICIAN_NAMES'].str.len()
    HOSP.list['PHYSICIAN_NAMES'] = HOSP.list['GEO_LOCATION'].map(doc_dict_new).fillna('')
    HOSP.list['NUM_PHYSICIANS'] = HOSP.list['PHYSICIAN_NAMES'].str.len()
    UPCC.list['PHYSICIAN_NAMES'] = UPCC.list['GEO_LOCATION'].map(doc_dict_new).fillna('')
    UPCC.list['NUM_PHYSICIANS'] = UPCC.list['PHYSICIAN_NAMES'].str.len()
    QFD.list['PHYSICIAN_NAMES'] = QFD.list['GEO_LOCATION'].map(doc_dict_new).fillna('')
    QFD.list['NUM_PHYSICIANS'] = QFD.list['PHYSICIAN_NAMES'].str.len()

def hotfixes():
    CPSBC.list.loc[(CPSBC.list['GYN'].str.contains(r"((\b|^)FAMILY\s*?AND\s*?MATERNITY(\b|$))",regex=True, na=False)), 'IS_FAMILY'] = 1


if __name__ == "__main__":
    # headings to be used in category separation
    headings = [
                'PO_BOX',
                'WEBSITE',
                'CORRECTIONS',
                'FIRST_NATIONS',
                'ADMIN',                
                'HOSP',
                'MHSU',
                'LTC',
                'DERM', 
                'REHAB', 
                'PLASTICS', 
                'PALLIATIVE', 
                'GYN', 
                'PAIN', 
                'GERIATRIC', 
                'IMAGE', 
                'SEX', 
                'ONCO', 
                'DISORDER', 
                'CNS', 
                'SURG', 
                'ORTHO', 
                'TRAVEL', 
                'VASCULAR',
                'VIRTUAL',
                'AGENCY',
                'UNI',
                'CLINIC',
                'CENTRE',
                'MED',
                'SOCIETY',
                'UNIT',
                'HEALTH',
                'FAMILY'
            ]
    
    categories = [
            'OTHER', 
            'CORRECTIONS',
            'FIRST_NATIONS',
            'ADMIN',                
            'HOSP',
            'MHSU',
            'LTC',
            'DERM', 
            'REHAB', 
            'PLASTICS', 
            'PALLIATIVE', 
            'GYN', 
            'PAIN', 
            'GERIATRIC', 
            'IMAGE', 
            'SEX', 
            'ONCO', 
            'DISORDER', 
            'CNS', 
            'SURG', 
            'ORTHO',  
            'TRAVEL', 
            'VASCULAR',
            'VIRTUAL',
            'AGENCY',
            'UNI',
            'CLINIC',
            'CENTRE',
            'MED',
            'SOCIETY',
            'UNIT',
            'HEALTH',
            'FAMILY',
            ]


    writer = pd.ExcelWriter('EXAMPLE_FINAL.xlsx', engine='openpyxl')
    CPSBC = Modified('materials\\EXAMPLE_INPUT_STEP2.xlsx', 'CPSBC Modified')
    WALKIN = Modified('materials\\EXAMPLE_INPUT_STEP2.xlsx', 'Walk-in Clinic List')
    CORR_FAC = Modified('materials\\EXAMPLE_INPUT_STEP2.xlsx', 'Corrections Facilities List')
    GSR = Modified('materials\\EXAMPLE_INPUT_STEP2.xlsx', 'GSR List')
    HOSP = Modified('materials\\EXAMPLE_INPUT_STEP2.xlsx', 'Hospitals List')
    UPCC = Modified('materials\\EXAMPLE_INPUT_STEP2.xlsx', 'UPCC List')
    QFD = Modified('materials\\EXAMPLE_INPUT_STEP2.xlsx', 'QFD List')
    CHSA = Modified('materials\\PC_CHSA.csv', sheetname=None)
    REGION = Modified('materials\\bc_health_region_master_2018.xlsx', 'CHSA')
    
    print('Geocoding Start')
    input_val = input('Please Enter Your Company Name (No Spaces): ')
    geolocator = DataBC(user_agent=input_val)
    geocoded = RateLimiter(geolocator.geocode, min_delay_seconds=1/15)
    print('Geocoding CPSBC List')
    CPSBC.geocode_flags(geolocator, geocoded, CPSBC_mod=True)
    print('Geocoding Walk-in Clinic List')
    WALKIN.geocode_flags(geolocator, geocoded)
    print('Geocoding Corrections Facilities List')
    CORR_FAC.geocode_flags(geolocator, geocoded)
    print('Geocoding GSR List')
    GSR.geocode_flags(geolocator, geocoded)
    print('Geocoding Hospital List')
    HOSP.geocode_flags(geolocator, geocoded)
    print('Geocoding UPCC List')
    UPCC.geocode_flags(geolocator, geocoded)
    print('Geocoding QFD List')
    QFD.geocode_flags(geolocator, geocoded)

    # Standarization and initialization of unique columns
    upper_it_all()

    # CPSBC standardization
    CPSBC.misc_addons(categories, renames={'ADDRESS FULL':'ADDRESS_FULL'}, empties=['WEBSITE'], CPSBC_mod=True)

    # Deals with PO Boxes that couldnt be geocoded
    CPSBC.po_fix()

    # CHSA and Region Standardization
    CHSA.list.rename(columns={'POSTALCODE': 'POSTAL', 'CHSA':'CHSA_NUM'}, inplace = True)
    CHSA.postal_standardization(upper=True)
    REGION.list.rename(columns={'CHSA_CD': 'CHSA_NUM'}, inplace = True)

    # UPCC standardization
    UPCC.misc_addons(categories, renames={'POSTAL_CODE':'POSTAL', 'PHONE_NUMBER':'PHONE'}, empties=['FAX']) 

    # GSR standardization
    GSR.misc_addons(categories, renames={'STREET_ADDRESS':'STREET_NUMBER', 'BUSINESS_PHONE':'PHONE', 'POSTAL_CODE': 'POSTAL', 'INSPECTION_URL':'WEBSITE'}, empties=['FAX']) 
    GSR.list['PHONE'] = GSR.list['PHONE'].str.replace(r'\D', '', regex=True)

    # QFD Standardization
    QFD.misc_addons(categories, renames={'STREET_ADDRESS':'STREET_NUMBER', 'WEBPAGE':'WEBSITE'}, empties=['FAX']) 
    QFD.list['PHONE'] = QFD.list['PHONE'].str.replace(r'\D', '', regex=True)
    QFD.list['POSTAL'] = QFD.list['POSTAL'].astype(str).apply(lambda x: (x[:3]+" "+x[3:]) if x != None else None)

    # CORRECTIONS Standardization
    CORR_FAC.misc_addons(categories, renames={'PHONE_NUMBER':'PHONE', 'POSTAL_CODE': 'POSTAL', 'FAX_NUMBER':'FAX'})

    # WALKINS Standardization
    WALKIN.misc_addons(categories, renames={'PHONE_NUMBER':'PHONE', 'POSTAL_CODE': 'POSTAL'}, empties=['FAX'])

    # HOSP Standardization
    HOSP.misc_addons(categories, renames={'PHONE_NUMBER':'PHONE', 'POSTAL_CODE': 'POSTAL'}, empties=['FAX'])

    # Preps and adds physicians to addresses
    CPSBC.find_duplicate_add()
    apply_physicians()

    # CPSBC Dictionaries
    CPSBC.list_application(categories)

    # Addresses from lists added, Order matters
    UPCC.list_application(categories, list_type='UPCC')
    GSR.list_application(categories, list_type='LTC')
    QFD.list_application(categories, list_type='LTC')
    CORR_FAC.list_application(categories, list_type='CORRECTIONS')
    WALKIN.list_application(categories, list_type='WALKIN')
    HOSP.list_application(categories, list_type='HOSP')

    # Dictionary Application
    CHSA.dict_application('CHSA_NUM','POSTAL')
    region_dicts = ['CHSA_NAME', 'HA_ID', 'HA_NAME', 'CHSA_UR_CLASS']
    for item in region_dicts:
        REGION.dict_application(item,'CHSA_NUM')
    for item in headings:
        if item == 'HOSP':
            CPSBC.dict_application(item, 'GEO_ADDRESS')
        else:
            CPSBC.dict_application(item, 'GEO_LOCATION')

    CPSBC.list.drop_duplicates(subset = ['GEO_LOCATION'], keep = 'last', inplace=True)

    # Applies Temporary Fixes (Saves time geocoding)
    hotfixes()
    assign_id(categories)

    CPSBC.out_sheet(writer, 'Clinic List Complete')
    # WALKIN.out_sheet(writer, 'Walkin')
    # CORR_FAC.out_sheet(writer, 'Corrfac')
    # GSR.out_sheet(writer, 'GSR')
    # HOSP.out_sheet(writer, 'Hospital')
    # UPCC.out_sheet(writer, 'UPCC')
    # QFD.out_sheet(writer, 'QFD')
    writer.save()









    

    
