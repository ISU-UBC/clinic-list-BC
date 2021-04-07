import numpy as np
import pandas as pd
import re
import camp

class DictOfDicts():
    def __init__(self, walkin_clinic = False):
        self.walkin_clinic = walkin_clinic
        self.po_dict = {}
        self.website_dict = {}
        self.corrections_dict = {}
        self.fn_dict = {}
        self.admin_dict = {}
        self.hospital_dict = {}
        self.mhsu_dict = {}
        self.ltc_dict = {}
        self.derm_dict = {} 
        self.rehab_dict = {} 
        self.plastics_dict = {} 
        self.palliative_dict = {} 
        self.gyn_dict = {} 
        self.pain_dict = {} 
        self.geriatric_dict = {} 
        self.image_dict = {} 
        self.sex_dict = {} 
        self.onco_dict = {} 
        self.disorder_dict = {} 
        self.cns_dict = {} 
        self.surg_dict = {} 
        self.ortho_dict = {} 
        self.travel_dict = {} 
        self.vascular_dict = {}
        self.virtual_dict = {}
        self.agency_dict = {}
        self.university_dict = {}
        self.clinic_dict = {}
        self.centre_dict = {}
        self.med_dict = {}
        self.society_dict = {}
        self.unit_dict = {}
        self.health_dict = {}
        self.family_med_dict = {}
        self.dictionaries_dict = {
                'PO_BOX_ADDR':self.po_dict,
                'WEBSITE':self.website_dict,
                'CORRECTIONS':self.corrections_dict,
                'FIRST_NATIONS':self.fn_dict,
                'ADMIN':self.admin_dict,                
                'HOSP':self.hospital_dict,
                'MHSU':self.mhsu_dict,
                'LTC':self.ltc_dict,
                'DERM':self.derm_dict, 
                'REHAB':self.rehab_dict, 
                'PLASTICS':self.plastics_dict, 
                'PALLIATIVE':self.palliative_dict, 
                'GYN':self.gyn_dict, 
                'PAIN':self.pain_dict, 
                'GERIATRIC':self.geriatric_dict, 
                'IMAGE':self.image_dict, 
                'SEX':self.sex_dict, 
                'ONCO':self.onco_dict, 
                'DISORDER':self.disorder_dict, 
                'CNS':self.cns_dict, 
                'SURG':self.surg_dict, 
                'ORTHO':self.ortho_dict, 
                'TRAVEL':self.travel_dict, 
                'VASCULAR':self.vascular_dict,
                'VIRTUAL':self.virtual_dict,
                'AGENCY':self.agency_dict,
                'UNI':self.university_dict,
                'CLINIC':self.clinic_dict,
                'CENTRE':self.centre_dict,
                'MED':self.med_dict,
                'SOCIETY':self.society_dict,
                'UNIT':self.unit_dict,
                'HEALTH':self.health_dict,
                'FAMILY':self.family_med_dict
                }

    def create_dict(self, df, address_col, location_col, index=False):
        for key,value in self.dictionaries_dict.items():
            key = key.replace('_1', '')
            key = key.replace('_2', '')
            if index != False:
                key = key+"_{}".format(index)
            if value == self.hospital_dict:
                value.update(pd.Series(df.loc[(df[key].notna())&(df[address_col].notna()), key].values,index=df.loc[(df[key].notna())&(df[address_col].notna()), address_col]).to_dict())
            else:
                value.update(pd.Series(df.loc[(df[key].notna())&(df[location_col].notna()), key].values,index=df.loc[(df[key].notna())&(df[location_col].notna()), location_col]).to_dict())
            value = {k: v for k, v in value.items() if pd.Series(v).notna().all()}       
        return df

    def apply_dict(self, df, address_col, location_col, index=False):
        df = df.reset_index(drop=True)
        for key, value in self.dictionaries_dict.items():
            key = key.replace('_1', '')
            key = key.replace('_2', '')
            temp = ''
            if index != False:
                key = key+"_{}".format(index)
                temp = temp+"_{}".format(index)
            if value == self.hospital_dict:
                df[key] = df[address_col].map(value).fillna(df[key])
            else:
                df[key] = df[location_col].map(value).fillna(df[key])
            if self.walkin_clinic != False:
                df.loc[df[address_col].isin(value.keys()), f'walkin_clinic{temp}']=1

        return df


def results_update(modified, results_dict, lazydict, filtername):
    results_dict['filter applied'].append(filtername)
    results_dict['total entries (rows)'].append(len(modified.index))
    temp = pd.concat([modified['Postal 1'], modified['Postal 2']], axis=0)
    results_dict['total addresses'].append(temp.count())
    results_dict['total unique postal codes'].append(temp.nunique())
    results_dict['total rural'].append(modified['is_rural'].sum())
    results_dict['total unknown_clin'].append(modified['unknown_clin'].sum())
    results_dict['total unique years_in_prac'].append(modified["years_in_prac"].nunique()) 
    results_dict['total walkin_clinic'].append(modified['walkin_clinic'].sum()) 
    results_dict['total multi_clin'].append(modified['multi_clin'].sum())
    results_dict['total CHSA_num'].append(modified[['CHSA_num_1','CHSA_num_2']].nunique().sum())
    results_dict['total CHSA_name'].append(modified[['CHSA_Name_1','CHSA_Name_2']].nunique().sum())
    results_dict['total health_auth'].append(modified[['HA_ID_1','HA_ID_2']].nunique().sum())
    results_dict['total websites'].append(modified[['WEBSITE_1','WEBSITE_2']].nunique().sum())
    results_dict['total grad_IMG'].append(modified['grad_IMG'].sum())

    # laziness
    lazydict_new = {}
    for key, value in lazydict.items():
        key = re.sub(r'_[12]', '', key)
        lazydict_new[key] = value 
    for key, value in lazydict_new.items():
        results_dict['total '+ key].append(modified[value].sum())


def remove_units(df, suite_col, address_col, location_col):
    df[suite_col] = df[address_col].str.extract(r'(.*)\b\s--\s', expand=False)
    df[address_col] = df[address_col].str.replace(r'(.*)(\b\s--\s)', '', regex=True)
    df[location_col] = df[location_col].str.replace(r'(^|.*?)(\D*)(\s--\s)', r'\1\3', regex=True)
    df[location_col] = df[location_col].str.replace(r'(^|\D*?\b\s)(\w*?)(\s--\s)', r'\2\3', regex=True)


def address_adjust(modified, results_dict, filtername):
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
        'multi_clin'
            ]

    for item in init:
        modified[item] = 0

# Checks if there are values in associated columns
    tracking = {}
    for i in range (1,3):
        tracking.update({f'PO_BOX_ADDR_{i}':'has_PO_box'})
        tracking.update({item.replace('is_', '')+'_{}'.format(i):item for item in init[1:-3]})
        for k,v in tracking.items():
            if i == 2:
                temp = k[:-1] + '1'
                modified.loc[~(modified[k].isna())&~(modified[k]=='')&~(modified[temp]==modified[k]), v] += 1
            else:
                modified.loc[~(modified[k].isna())&~(modified[k]==''), v] += 1
        
        # Checking if unknown clinic
        check = [item.replace('is_', '')+'_{}'.format(i) for item in init[2:-3]]
        modified[f'temp_{i}'] = 0
        for item in check:
            modified.loc[~(modified[item].isna())&~(modified[item]=='')&~(modified[f'GEO_ADDRESS_{i}']=='')&~(modified[f'GEO_ADDRESS_{i}'].isna()), f'temp_{i}'] += 1
        modified.loc[(modified[f'temp_{i}']==0)&~(modified[f'GEO_ADDRESS_{i}']=='')&~(modified[f'GEO_ADDRESS_{i}'].isna()), 'unknown_clin'] += 1
        modified.drop(columns = [f'temp_{i}'], inplace=True)

    # Applying if doc goes to multiple clinic
    modified.loc[modified['is_CLINIC'].astype(int) > 1, 'multi_clin'] = 1

    # Applying if address is is_rural
    modified.loc[modified['Postal 1'].str.get(1)=='0', 'is_rural'] += 1
    modified.loc[(modified['Postal 2'].str.get(1)=='0')&~(modified['Postal 1']==modified['Postal 2']), 'is_rural'] += 1

    # Update results    
    results_update(modified, results_dict, tracking, filtername)


if __name__ == "__main__":
    # Opens the file (must be deencrypted)
    modified = pd.read_excel('working_list.xlsx', sheet_name='Modified Final')
    corrfac = pd.read_excel('working_list.xlsx', sheet_name='Corrections Facilities Final')
    hlbc = pd.read_excel('working_list.xlsx', sheet_name='HLBC Final')
    chsa = pd.read_csv('PC_CHSA.csv')
    ha = pd.read_excel('bc_health_region_master_2018.xlsx', sheet_name='CHSA')
    writer = pd.ExcelWriter('filtered_list.xlsx', engine='openpyxl')
    modified.rename(columns={'PO_BOX_ADDR 1': 'PO_BOX_ADDR_1', 'PO_BOX_ADDR 2': 'PO_BOX_ADDR_2'}, inplace = True)
    corrfac.rename(columns={'MAIL_ADDRESS': 'PO_BOX_ADDR'}, inplace = True)
    hlbc['PO_BOX_ADDR'] = np.nan
    # Results header names
    results_dict = {'filter applied':[], 
                    'total entries (rows)':[], 
                    'total unique postal codes':[],
                    'total addresses':[],
                    'total rural':[], 
                    'total unknown_clin':[],
                    'total unique years_in_prac':[], 
                    'total walkin_clinic':[], 
                    'total multi_clin':[], 
                    'total CHSA_num':[],
                    'total CHSA_name':[],
                    'total health_auth':[], 
                    'total websites':[],
                    'total grad_IMG':[],
                    'total PO_BOX_ADDR':[],
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
                    }

    # CHSA addition
    # Makes sure the Postal Codes all adhere to the same format
    chsa['POSTALCODE'] = chsa['POSTALCODE'].str.replace(' ', '')
    chsa['POSTALCODE'] = chsa['POSTALCODE'].apply(lambda x: (x[:3]+" "+x[3:]) if x != None else None)  

    # Creates CHSA and HA dictionary
    chsa_dict = {}
    chsa_dict.update(pd.Series(chsa['CHSA'].values,index=chsa['POSTALCODE']).to_dict())
    ha_list = ['CHSA_Name', 'HA_ID', 'HA_Name', 'CHSA_UR_Class']

    modified['walkin_clinic'] = 0

    # Dictionary creation
    modified_dict = DictOfDicts()
    # Remove commas in the unspecified columns
    for i in range (1,3):
        modified[f'WEBSITE_{i}'] = np.nan
        modified[f'CHSA_num_{i}'] = np.nan
        modified[f'walkin_clinic_{i}'] = 0
    # Separates units/suites/floors, etc from geocoded addresses
        remove_units(modified, f'GEO_SUITE_{i}', f'GEO_ADDRESS_{i}', f'GEO_LOCATION_{i}')

    # Creates dictionaries from the modified sheet itself
        modified = modified_dict.create_dict(modified, f'GEO_ADDRESS_{i}',f'GEO_LOCATION_{i}', index=i)
        
    # Applies CHSA dictionary to clinic list
        modified[f'CHSA_num_{i}'] = modified[f'Postal {i}'].map(chsa_dict).fillna(modified[f'CHSA_num_{i}'])
        for item in ha_list:
            ha_dict = {}
            ha_dict.update(pd.Series(ha[item].values,index=ha['CHSA_CD']).to_dict())
            modified[item+"_{}".format(i)] = modified[f'CHSA_num_{i}'].map(ha_dict).fillna(np.nan)

    # Separates units/suites/floors, etc from geocoded addresses
    remove_units(corrfac, 'GEO_SUITE', 'GEO_ADDRESS', 'GEO_LOCATION')
    remove_units(hlbc, 'GEO_SUITE', 'GEO_ADDRESS', 'GEO_LOCATION')

    
    # Baseline
    address_adjust(modified, results_dict, filtername='Baseline')


    # Filter 1: Specialty
    modified = modified[modified['Specialties & Certifications'].str.contains('^Fam.*', regex=True)==True|modified['Specialties & Certifications'].str.contains('NaN', regex=False)|modified['Specialties & Certifications'].str.contains(None, regex=False)]
    address_adjust(modified, results_dict, filtername='Family Docs')


    # Filter 2: BC
    modified = modified[(modified['Province 1']=='BC')|(modified['Province 2']=='BC')]
    address_adjust(modified, results_dict, filtername='BC')


    # Dealing with the PO BOXES
    for i in range (1,3):
        address = [f'PO_BOX_ADDR_{i}', f'City {i}', f'Province {i}', f'Postal {i}']
        modified.loc[~(modified[f'GEO_ADDRESS_{i}'].str.contains(r"(\d+)",regex=True, na=False))&(modified[f'Province {i}']=='BC'), [f'GEO_ADDRESS_{i}',f'GEO_LOCATION_{i}']] = modified[address].apply(lambda x: ', '.join(x.dropna()), axis=1) 
        camp.ditch_commas(modified, f'GEO_ADDRESS_{i}')
        camp.ditch_commas(modified, f'GEO_LOCATION_{i}')

    # Filter 3: Initial Dictionaries  
    # Applies dictionaries to clinic list
    for i in range (1,3):
        modified = modified_dict.apply_dict(modified, f'GEO_ADDRESS_{i}',f'GEO_LOCATION_{i}', index=i) 
    address_adjust(modified, results_dict, filtername='Dictionaries')

    modified.to_excel(writer, sheet_name = 'Filters 1-5', index=False)

    # Results to a dataframe for easy export
    results = pd.DataFrame.from_dict(results_dict)

    # Final Exportation
    results.to_excel(writer, sheet_name = 'Result of Filters', index=False)
    writer.save()








    

    
