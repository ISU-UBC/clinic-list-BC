import numpy as np
import pandas as pd
import re
import camp
import filters


def assign_id(clin_list_final, working_list, boolean, droplist, id_name, sheet_name):
    temp = working_list.loc[boolean]
    working_list = working_list.loc[~boolean]
    temp.reset_index(drop=True, inplace=True)
    working_list.reset_index(drop=True, inplace=True)
    temp['ID'] = 1 + temp.index
    temp['ID'] = temp['ID'].apply('{:0>4}'.format)
    temp['ID'] = id_name + temp['ID'].astype(str)
    temp2 = temp.drop(columns = droplist)
    temp2.to_excel(writer, sheet_name = sheet_name, index=False)
    clin_list_final = pd.concat([clin_list_final, temp], axis=0)

    return clin_list_final, working_list

def results_update(modified, results_dict, lazydict):
    results_dict['total entries (rows)'].append(len(modified.index))
    results_dict['total unique postal codes'].append(modified['POSTAL_CODE'].nunique())
    results_dict['total rural'].append(modified['is_rural'].sum())
    results_dict['total unknown_clin'].append(modified['unknown_clin'].sum())
    results_dict['total walkin_clinic'].append(modified['WALKIN_CLINIC'].sum()) 
    results_dict['total UPCC'].append(modified['UPCC_LIST'].sum()) 
    results_dict['total multi_practitioner'].append(modified['multi_practitioner'].sum())
    results_dict['total CHSA_NUM'].append(modified['CHSA_NUM'].nunique())
    results_dict['total CHSA_NAME'].append(modified['CHSA_NAME'].nunique())
    results_dict['total health_auth'].append(modified['HA_ID'].nunique())
    results_dict['total websites'].append(modified['WEBSITE'].nunique())

    # laziness
    lazydict_new = {}
    for key, value in lazydict.items():
        lazydict_new[key] = value 
    for key, value in lazydict_new.items():
        results_dict['total '+ key].append(modified[value].sum())

def remove_units(df, suite_col, address_col):
    df[suite_col] = df[address_col].str.extract(r'(.*)\b\s--\s', expand=False)
    df[address_col] = df[address_col].str.replace(r'(.*)(\b\s--\s)', '', regex=True)

def addons_ops(df, tempheadings, headings, ha_list):
    df['BRACKETS'] = np.nan
    df['COUNTRY'] = 'CANADA'
    df['ADDRESS_FULL'] = df[tempheadings].apply(lambda x: ', '.join(x.dropna()), axis=1)
    camp.ditch_commas(df, 'ADDRESS_FULL')
    # Formatting for join later
    df = df[headings]
    for item in ha_list:
        df[item] = ''
    return df


def address_adjust(modified, results_dict):
# Initializes results columns with 0
    init = [
        # Tracking portion
        'has_PO_box', 
        'is_OTHER', 
        # Used to check for unknown clinic
        'is_CORRECTIONS',
        'is_FIRST_NATIONS',
        'is_ADMIN',                
        'is_HOSP',
        'is_MHSU',
        'is_LTC',
        'is_DERM', 
        'is_REHAB', 
        'is_PLASTICS', 
        'is_PALLIATIVE', 
        'is_GYN', 
        'is_PAIN', 
        'is_GERIATRIC', 
        'is_IMAGE', 
        'is_SEX', 
        'is_ONCO', 
        'is_DISORDER', 
        'is_CNS', 
        'is_SURG', 
        'is_ORTHO', 
        'is_TRAVEL', 
        'is_VASCULAR',
        'is_VIRTUAL',
        'is_AGENCY',
        'is_UNI',
        'is_CLINIC',
        'is_CENTRE',
        'is_MED',
        'is_SOCIETY',
        'is_UNIT',
        'is_HEALTH',
        'is_FAMILY',
        # End tracking portion unknown clinic
        'unknown_clin',
        # after check for unknown clinic
        'is_rural',
        'multi_practitioner'
            ]

    for item in init:
        modified[item] = 0

# Checks if there are values in associated columns
    tracking = {}
    tracking.update({'PO_BOX':'has_PO_box'})
    tracking.update({item.replace('is_', ''):item for item in init[1:-3]})
    for k,v in tracking.items():
        modified.loc[~(modified[k].isna())&~(modified[k]==''), v] += 1
    
    # Checking if unknown clinic
    check = [item.replace('is_', '') for item in init[2:-3]]
    modified['temp'] = 0
    for item in check:
        modified.loc[~(modified[item].isna())&~(modified[item]=='')&~(modified['GEO_LOCATION']=='')&~(modified['GEO_LOCATION'].isna()), 'temp'] += 1
    modified.loc[(modified['temp']==0)&~(modified['GEO_LOCATION']=='')&~(modified['GEO_LOCATION'].isna()), 'unknown_clin'] += 1
    modified.drop(columns = ['temp'], inplace=True)

    # Applying if doc goes to multiple clinic
    modified.loc[modified['NUM_PHYSICIANS'].astype(int) > 1, 'multi_practitioner'] = 1

    # Applying if address is is_rural
    modified.loc[modified['POSTAL_CODE'].str.get(1)=='0', 'is_rural'] += 1

    # Special Corrections Facility Case(s) 
    modified.loc[(modified['GEO_LOCATION']=='70 Colony Farm Rd, Coquitlam, BC'), 'is_CORRECTIONS'] = 1

    # Quick Fixes for custom searches go here 
    modified.loc[(modified['UPCC_LIST']==1), 'UPCC'] = 1
    modified.loc[(modified['LTC_LIST']==1), 'is_LTC'] = 1
    modified.loc[(modified['CORRECTIONS_LIST']==1), 'is_CORRECTIONS'] = 1
    modified.loc[(modified['WALKIN_CLINIC_LIST']==1), 'WALKIN_CLINIC'] = 1  
    modified.loc[(modified['HOSP_LIST']==1), 'is_HOSP'] = 1

    # Update results    
    results_update(modified, results_dict, tracking)
    return modified

if __name__ == "__main__":
    # Opens the file (must be deencrypted)
    modified = pd.read_excel('filtered_list.xlsx', sheet_name='Filters 1-5').fillna('')
    corrfac = pd.read_excel('working_list.xlsx', sheet_name='Corrections Facilities Final').fillna('')
    hlbc = pd.read_excel('working_list.xlsx', sheet_name='HLBC Final').fillna('')
    chsa = pd.read_csv('PC_CHSA.csv').fillna('')
    region_master = pd.read_excel('bc_health_region_master_2018.xlsx', sheet_name='CHSA').fillna('')
    region_master.columns = map(str.upper, region_master.columns)
    region_master.rename(columns={'CHSA_CD':'CHSA_NUM'}, inplace=True)
    hosp = pd.read_excel('flagged_added.xlsx', sheet_name = 'hosp_geo').fillna('')
    upcc = pd.read_excel('flagged_added.xlsx', sheet_name = 'upcc_geo').fillna('')
    gsr = pd.read_excel('flagged_added.xlsx', sheet_name = 'gsr_geo').fillna('')
    qfd = pd.read_excel('flagged_added.xlsx', sheet_name = 'qfd_geo').fillna('')
    writer = pd.ExcelWriter('clinic_list.xlsx', engine='openpyxl')

    # CHSA addition to hlbc & corrections
    # Makes sure the Postal Codes all adhere to the same format
    chsa['POSTALCODE'] = chsa['POSTALCODE'].str.replace(' ', '')
    chsa['POSTALCODE'] = chsa['POSTALCODE'].apply(lambda x: (x[:3]+" "+x[3:]) if x != None else None)  
    # Creates CHSA dictionary
    chsa_dict = {}
    chsa_dict.update(pd.Series(chsa['CHSA'].values,index=chsa['POSTALCODE']).to_dict())
    chsa_df = pd.DataFrame(chsa_dict, index=[0])
   
    #list of headings to use    
    headings = [
                'PO_BOX',
                'ADDRESS_FULL',
                'CITY',
                'PROVINCE',
                'POSTAL_CODE',
                'COUNTRY',
                'PHONE_NUMBER', 
                'FAX_NUMBER',
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
                'GEO_LOCATION',
                'GEO_ADDRESS', 
                'GEO_LATITUDE', 
                'GEO_LONGITUDE', 
                'WALKIN_CLINIC',
                'UPCC',
                'WALKIN_CLINIC_LIST', 
                'CORRECTIONS_LIST', 
                'UPCC_LIST', 
                'HOSP_LIST', 
                'LTC_LIST'
                ]
    
    tempheadings = headings[8:-13]
    tempheadings.append('STREET_NUMBER')
    
    # TEST CASES
    # addons = [hlbc, corrfac]
    # addons = [hlbc, corrfac, hosp]
    # addons = [hlbc, corrfac, hosp, upcc]
    # addons = [hlbc, corrfac, hosp, upcc, gsr]
    addons = [upcc, gsr, qfd, corrfac, hlbc, hosp]
    
    ha_list = ['CHSA_NAME', 'HA_ID', 'HA_NAME', 'CHSA_UR_CLASS', 'CHSA_NUM']
    # Some Cleanup
    added_list_cols = [
        'UPCC',
        'WALKIN_CLINIC',
        'WALKIN_CLINIC_LIST', 
        'CORRECTIONS_LIST', 
        'UPCC_LIST', 
        'HOSP_LIST', 
        'LTC_LIST'
        ]
    for item in addons:
        for column in added_list_cols:
            item[column] = 0

    # UPCC
    upcc['FAX_NUMBER'] = np.nan
    upcc['PO_BOX'] = np.nan
    upcc['UPCC_LIST'] = 1
    upcc['PO_BOX'] = upcc['ADDR_FOR_GEO'].str.extract(r'([PO]*\s?BOX\w?.*?\d+|BAG*\s*?\d+)', expand=False)
    upcc['ADDR_FOR_GEO'] = upcc['ADDR_FOR_GEO'].str.replace(r'([PO]*\s?BOX\w?.*?\d+|BAG*\s*?\d+)', '', regex=True) 
    upcc['PO_BOX'] = upcc['PO_BOX'].str.upper()
    filters.remove_units(upcc, 'GEO_SUITE', 'GEO_ADDRESS', 'GEO_LOCATION')
    upcc = addons_ops(upcc, tempheadings, headings, ha_list)

    # GSR
    gsr.rename(columns={'STREET_ADDRESS':'STREET_NUMBER', 'BUSINESS_PHONE':'PHONE_NUMBER', 'INSPECTION_URL':'WEBSITE'}, inplace=True)
    gsr['PHONE_NUMBER'] = gsr['PHONE_NUMBER'].str.replace(r'\D', '', regex=True)
    gsr['FAX_NUMBER'] = np.nan
    gsr['PO_BOX'] = np.nan
    gsr['LTC_LIST'] = 1
    gsr['PO_BOX'] = gsr['ADDR_FOR_GEO'].str.extract(r'([PO]*\s?BOX\w?.*?\d+|BAG*\s*?\d+)', expand=False)
    gsr['ADDR_FOR_GEO'] = gsr['ADDR_FOR_GEO'].str.replace(r'([PO]*\s?BOX\w?.*?\d+|BAG*\s*?\d+)', '', regex=True) 
    gsr['PO_BOX'] = gsr['PO_BOX'].str.upper()
    filters.remove_units(gsr, 'GEO_SUITE', 'GEO_ADDRESS', 'GEO_LOCATION')
    gsr = addons_ops(gsr, tempheadings, headings, ha_list)

    # QFD
    qfd.rename(columns={'STREET_ADDRESS':'STREET_NUMBER', 'PHONE':'PHONE_NUMBER', 'WEBPAGE':'WEBSITE', 'POSTAL':'POSTAL_CODE'}, inplace=True)
    qfd['PHONE_NUMBER'] = qfd['PHONE_NUMBER'].str.replace(r'\D', '', regex=True)
    qfd['FAX_NUMBER'] = np.nan
    qfd['PO_BOX'] = np.nan
    qfd['LTC_LIST'] = 1
    qfd['PO_BOX'] = qfd['ADDR_FOR_GEO'].str.extract(r'([PO]*\s?BOX\w?.*?\d+|BAG*\s*?\d+)', expand=False)
    qfd['ADDR_FOR_GEO'] = qfd['ADDR_FOR_GEO'].str.replace(r'([PO]*\s?BOX\w?.*?\d+|BAG*\s*?\d+)', '', regex=True) 
    qfd['PO_BOX'] = qfd['PO_BOX'].str.upper()
    qfd['POSTAL_CODE'] = qfd['POSTAL_CODE'].astype(str).apply(lambda x: (x[:3]+" "+x[3:]) if x != None else None)
    filters.remove_units(qfd, 'GEO_SUITE', 'GEO_ADDRESS', 'GEO_LOCATION')
    qfd = addons_ops(qfd, tempheadings, headings, ha_list)

    # CORRECTIONS
    corrfac['CORRECTIONS_LIST'] = 1
    corrfac.rename(columns={'MAIL_ADDRESS':'PO_BOX'}, inplace=True)
    corrfac.loc[~(corrfac['PO_BOX'].str.contains(r"([\b|^]Bag\b|\bBox\b)",regex=True, na=True)), 'PO_BOX'] = np.nan
    filters.remove_units(corrfac, 'GEO_SUITE', 'GEO_ADDRESS', 'GEO_LOCATION')
    corrfac = addons_ops(corrfac, tempheadings, headings, ha_list)

    # WALKINS
    hlbc['FAX_NUMBER'] = np.nan
    hlbc['PO_BOX'] = np.nan
    hlbc['WALKIN_CLINIC_LIST'] = 1
    filters.remove_units(hlbc, 'GEO_SUITE', 'GEO_ADDRESS', 'GEO_LOCATION')
    hlbc = addons_ops(hlbc, tempheadings, headings, ha_list)

    # HOSP
    hosp['HOSP_LIST'] = 1
    hosp['PO_BOX'] = hosp['ADDR_FOR_GEO'].str.extract(r'([PO]*\s?BOX\w?.*?\d+|BAG*\s*?\d+)', expand=False)
    hosp['ADDR_FOR_GEO'] = hosp['ADDR_FOR_GEO'].str.replace(r'([PO]*\s?BOX\w?.*?\d+|BAG*\s*?\d+)', '', regex=True) 
    hosp['PO_BOX'] = hosp['PO_BOX'].str.upper()
    hosp['FAX_NUMBER'] = np.nan
    hosp = addons_ops(hosp, tempheadings, headings, ha_list)

     # Final Clinic list creation
    oldheadings = [
                    'PO_BOX_ADDR_1', 
                    'Address 1 Full',
                    'City 1', 
                    'Province 1', 
                    'Postal 1', 
                    'Country 1', 
                    'Phone 1', 
                    'Fax 1', 
                    'OTHER_1', 
                    'CORRECTIONS_1',
                    'FIRST_NATIONS_1',
                    'ADMIN_1',                
                    'HOSP_1',
                    'MHSU_1',
                    'LTC_1',
                    'DERM_1', 
                    'REHAB_1', 
                    'PLASTICS_1', 
                    'PALLIATIVE_1', 
                    'GYN_1', 
                    'PAIN_1', 
                    'GERIATRIC_1', 
                    'IMAGE_1', 
                    'SEX_1', 
                    'ONCO_1', 
                    'DISORDER_1', 
                    'CNS_1', 
                    'SURG_1', 
                    'ORTHO_1', 
                    'TRAVEL_1', 
                    'VASCULAR_1',
                    'VIRTUAL_1',
                    'AGENCY_1',
                    'UNI_1',
                    'CLINIC_1',
                    'CENTRE_1',
                    'MED_1',
                    'SOCIETY_1',
                    'UNIT_1',
                    'HEALTH_1',
                    'FAMILY_1',
                    'BRACKETS_1',
                    'WEBSITE_1',
                    'GEO_LOCATION_1',
                    'GEO_ADDRESS_1', 
                    'GEO_LATITUDE_1', 
                    'GEO_LONGITUDE_1', 
                    'CHSA_num_1', 
                    'CHSA_Name_1',
                    'CHSA_UR_Class_1',
                    'HA_ID_1',
                    'HA_Name_1',
                    'walkin_clinic_1'
                    ]
    
    clin_list_1 = modified[oldheadings]

    oldheadings = [item.replace('1', '2') for item in oldheadings]
    clin_list_2 = modified[oldheadings]

    temp = [clin_list_1, clin_list_2]
    for item in temp:
        item.columns = item.columns.str.replace(r'.\d', '')
        item.columns = item.columns.str.replace(r'\s', '_')
        item.columns = map(str.upper, item.columns)
        item.rename(columns={'PO_BOX_ADDR':'PO_BOX', 'POSTAL':'POSTAL_CODE', 'PHONE':'PHONE_NUMBER','FAX':'FAX_NUMBER'}, inplace=True)

    clin_list_final = pd.concat([clin_list_1, clin_list_2]).dropna(how='all')
    clin_list_final = clin_list_final[(clin_list_final['PROVINCE']=='BC')]

    # Dropping the duplicates based on full geo locations
    clin_list_final.drop_duplicates(subset=['GEO_LOCATION'], inplace=True, keep='last')
    clin_list_final[['UPCC', 'UPCC_LIST', 'WALKIN_CLINIC_LIST', 'CORRECTIONS_LIST', 'LTC_LIST', 'HOSP_LIST']] = 0

    addons = [upcc, gsr, qfd, corrfac, hlbc, hosp]    
    for item in addons:
        clin_list_final = pd.concat([clin_list_final, item]).dropna(how='all')
        clin_list_final.drop_duplicates(subset=['GEO_LOCATION'], inplace=True, keep='last')
    
    # Dropping the duplicates based on full geo locations
    clin_list_final.drop_duplicates(subset=['GEO_LOCATION'], inplace=True, keep='last')

    # Applies CHSA dictionary
    clin_list_final['CHSA_NUM'] = clin_list_final['POSTAL_CODE'].map(chsa_dict).fillna(np.nan)
    for some in ha_list:
        ha_dict = {}
        ha_dict.update(pd.Series(region_master[some].values,index=region_master['CHSA_NUM']).to_dict())
        clin_list_final[some.upper()] = clin_list_final['CHSA_NUM'].map(ha_dict).fillna(np.nan)

    # Instantiates empty dictionaries as they'll be updated in the loop below, applies docs to the clinics

    clin_list_final['PHYSICIAN_NAME'] = ''
    modified['DOC_NAME'] = modified[['Last Name', 'Given Names']].apply(lambda x: ', '.join(x.dropna()), axis=1)
    doc_dict_new = {}
    for i in range (1,3):
        temp = modified[modified[f'Province {i}']=='BC']
        doc_dict = {}
        doc_dict.update(pd.Series(temp[f'GEO_LOCATION_{i}'].values,index=temp['DOC_NAME']).to_dict())
        # keys = docs, values = geo_location
        for key, value in doc_dict.items():
            doc_dict_new[value] = []

    for i in range (1,3):
        temp = modified[modified[f'Province {i}']=='BC']
        doc_dict = {}
        doc_dict.update(pd.Series(temp[f'GEO_LOCATION_{i}'].values,index=temp['DOC_NAME']).to_dict())
        for key, value in doc_dict.items():
            if i == 2:
                if key + ' - 1' in doc_dict_new[value]:
                    doc_dict_new[value].remove(key + ' - 1')
                    doc_dict_new[value].extend([key + ' - 1&2'])
                else:
                    doc_dict_new[value].extend([key + f' - {i}']) 
            else:     
                doc_dict_new[value].extend([key + f' - {i}'])

    clin_list_final['PHYSICIAN_NAME'] = clin_list_final['GEO_LOCATION'].map(doc_dict_new).fillna(clin_list_final['PHYSICIAN_NAME'])
    clin_list_final.loc[~(clin_list_final['PHYSICIAN_NAME'].isna()), 'NUM_PHYSICIANS'] = clin_list_final['PHYSICIAN_NAME'].str.len()

    # Results header names
    results_dict = {
                    'total entries (rows)':[], 
                    'total unique postal codes':[],
                    'total rural':[], 
                    'total unknown_clin':[],
                    'total walkin_clinic':[], 
                    'total multi_practitioner':[], 
                    'total CHSA_NUM':[],
                    'total CHSA_NAME':[],
                    'total health_auth':[], 
                    'total websites':[],
                    'total PO_BOX':[],
                    'total OTHER':[], 
                    'total CORRECTIONS':[],
                    'total FIRST_NATIONS':[],
                    'total ADMIN':[],                
                    'total HOSP':[],
                    'total MHSU':[],
                    'total LTC':[],
                    'total DERM':[], 
                    'total REHAB':[], 
                    'total PLASTICS':[], 
                    'total PALLIATIVE':[], 
                    'total GYN':[], 
                    'total PAIN':[], 
                    'total GERIATRIC':[], 
                    'total IMAGE':[], 
                    'total SEX':[], 
                    'total ONCO':[], 
                    'total DISORDER':[], 
                    'total CNS':[], 
                    'total SURG':[], 
                    'total ORTHO':[], 
                    'total TRAVEL':[], 
                    'total VASCULAR':[],
                    'total VIRTUAL':[],
                    'total AGENCY':[],
                    'total UNI':[],
                    'total CLINIC':[],
                    'total CENTRE':[],
                    'total MED':[],
                    'total SOCIETY':[],
                    'total UNIT':[],
                    'total HEALTH':[],
                    'total FAMILY':[],
                    'total UPCC':[],
                    }

    # Results tabulations
    modified = address_adjust(clin_list_final, results_dict)
    

    # Sorts by geo address without unit then by Physician number (Requested by Rita)
    clin_list_final.sort_values(by=['GEO_ADDRESS','NUM_PHYSICIANS'], inplace=True)
    clin_list_final['ID'] = np.nan

    # Splits into separate categories
    # droplist = []
    droplist = [
                'WALKIN_CLINIC',
                'UPCC',
                # 'WALKIN_CLINIC_LIST', 
                # 'CORRECTIONS_LIST', 
                # 'UPCC_LIST', 
                # 'HOSP_LIST', 
                # 'LTC_LIST',
                'has_PO_box',
                'is_OTHER',
                'is_CORRECTIONS',
                'is_FIRST_NATIONS',
                'is_ADMIN',
                'is_HOSP',
                'is_MHSU',
                'is_LTC',
                'is_DERM',
                'is_REHAB',
                'is_PLASTICS',
                'is_PALLIATIVE',
                'is_GYN',
                'is_PAIN',
                'is_GERIATRIC',
                'is_IMAGE',
                'is_SEX',
                'is_ONCO',
                'is_DISORDER',
                'is_CNS',
                'is_SURG',
                'is_ORTHO',
                'is_TRAVEL',
                'is_VASCULAR',
                'is_VIRTUAL',
                'is_AGENCY',
                'is_UNI',
                'is_CLINIC',
                'is_CENTRE',
                'is_MED',
                'is_SOCIETY',
                'is_UNIT',
                'is_HEALTH',
                'is_FAMILY',
                'unknown_clin',
                'is_rural',
                'multi_practitioner'
                ]
                
    working_list = clin_list_final
    clin_list_final, working_list = assign_id(clin_list_final, working_list, boolean=(working_list['is_HOSP']==1), droplist = droplist, id_name='HOS_', sheet_name='Hospital ID') 
    clin_list_final, working_list = assign_id(clin_list_final, working_list, boolean=(working_list['WALKIN_CLINIC']==1), droplist = droplist, id_name='WIC_', sheet_name='Walk-in Clinic ID') 
    clin_list_final, working_list = assign_id(clin_list_final, working_list, boolean=(working_list['is_LTC']==1), droplist = droplist, id_name='LTC_', sheet_name='Long Term Care ID')
    clin_list_final, working_list = assign_id(clin_list_final, working_list, boolean=(working_list['UPCC']==1), droplist = droplist, id_name='UPC_', sheet_name='Urgent Primary Care Clinic ID')
    clin_list_final, working_list = assign_id(clin_list_final, working_list, boolean=(working_list['is_FAMILY']==1), droplist = droplist, id_name='FAM_', sheet_name='Family ID') 
    clin_list_final, working_list = assign_id(clin_list_final, working_list, boolean=((working_list['is_CORRECTIONS']==1)), droplist = droplist, id_name='CFI_', sheet_name='Corrections ID')
    clin_list_final, working_list = assign_id(clin_list_final, working_list, boolean=(working_list['is_FIRST_NATIONS']==1), droplist = droplist, id_name='FN_', sheet_name='First Nations ID')
    clin_list_final, working_list = assign_id(clin_list_final, working_list, boolean=(working_list['is_SEX']==1), droplist = droplist, id_name='SXH_', sheet_name='Sexual Health Clinic ID') 
    clin_list_final, working_list = assign_id(clin_list_final, working_list, boolean=(working_list['is_GYN']==1), droplist = droplist, id_name='GYN_', sheet_name='Womens Health Clinic ID')
    clin_list_final, working_list = assign_id(clin_list_final, working_list, boolean=(working_list['is_VIRTUAL']==1), droplist = droplist, id_name='VIR_', sheet_name='Virtual ID') 
    clin_list_final, working_list = assign_id(clin_list_final, working_list, boolean=(working_list['is_ADMIN']==1), droplist = droplist, id_name='ADM_', sheet_name='Administrative ID')
    specialty_bool = (working_list['is_MHSU']==1)|(working_list['is_DERM']==1)|(working_list['is_REHAB']==1)|(working_list['is_PLASTICS']==1)|(working_list['is_PALLIATIVE']==1)|(working_list['is_PAIN']==1)|(working_list['is_GERIATRIC']==1)|(working_list['is_IMAGE']==1)|(working_list['is_ONCO']==1)|(working_list['is_DISORDER']==1)|(working_list['is_CNS']==1)|(working_list['is_SURG']==1)|(working_list['is_ORTHO']==1)|(working_list['is_VASCULAR']==1)|(working_list['is_TRAVEL']==1)
    clin_list_final, working_list = assign_id(clin_list_final, working_list, boolean=(specialty_bool), droplist = droplist, id_name='SPC_', sheet_name='Specialty Clinic ID') 
    clin_list_final, working_list = assign_id(clin_list_final, working_list, boolean=((working_list['is_CLINIC']==1)|(working_list['is_CENTRE']==1)), droplist = droplist, id_name='PCC_', sheet_name='Clinic or Centre Not Walk-in ID')
    clin_list_final, working_list = assign_id(clin_list_final, working_list, boolean=(working_list['NUM_PHYSICIANS']>1), droplist = droplist, id_name='UNM_', sheet_name='Unknown Multi Prac ID')
    working_list['ID'] = 1 + working_list.index
    working_list['ID'] = working_list['ID'].apply('{:0>4}'.format)
    working_list['ID'] = 'UNS_' + working_list['ID'].astype(str)
    working_list.drop(columns = droplist, inplace = True)
    working_list.to_excel(writer, sheet_name = 'Unknown Single Prac ID', index=False)
    clin_list_final = pd.concat([clin_list_final, working_list], axis=0)
    clin_list_final.drop(columns = droplist, inplace = True)
    clin_list_final.dropna(subset=['ID'], inplace=True)

    # Results to a dataframe for easy export
    results = pd.DataFrame.from_dict(results_dict)

    # Final Exportation
    results.to_excel(writer, sheet_name = 'Clinic List Summary', index=False)
    clin_list_final.to_excel(writer, sheet_name = 'Clinic List', index=False)
    writer.save()









    

    
