# Code written by Cameron Lindsay
# Project Authors: Rita McCracken, Ian Cooper, Cameron Lindsay, Tiffany Hill
# For the CAMP/MAAP Clinic List Project

import numpy as np
import pandas as pd
from geopy.geocoders import DataBC
from geopy.extra.rate_limiter import RateLimiter
import geopy
import re

class List():

    def __init__(self, filename, sheetname=None, CPSBC_mod = False):
        self.filename = filename
        if CPSBC_mod == True:
            self.list = pd.read_excel(filename, sheet_name=sheetname, usecols='A:AJ').fillna('')
        elif sheetname==None:
            self.list = pd.read_csv(filename).fillna('')
        else:
            self.list = pd.read_excel(filename, sheet_name=sheetname).fillna('')
    
    def out_sheet(self, writer, sheetname):
        self.list.to_excel(writer, sheet_name = sheetname, index=False)

    def drop_and_combine(self):
         # Removes non-family med specialists and addresses outside of BC
        def formatting(columns, index):
            df = self.list[columns].dropna(subset=[f'Province {index}'])
            df = df[(df[f'Province {index}']=='BC')]
            df['add_listing'] = index
            df.rename(columns={f'Address {index} Line 1': 'Address Line 1'}, inplace=True)
            df.rename(columns={f'Address {index} Line 2': 'Address Line 2'}, inplace=True)
            for item in columns:
                if item != 'Address Line 1' and item != 'Address Line 2':
                    temp = re.sub(r'\d', '', item)
                    temp = temp.strip()           
                    df.rename(columns={item: temp}, inplace=True)       
            return df

        columns = self.list.columns.tolist()
        common = ['Last Name', 'Given Names']
        address1 = list(np.append(common, columns[16:-10]))
        address2 = list(np.append(common, columns[-10:]))
        self.list = self.list[(self.list['Specialties & Certifications'].str.contains(r'^Fam.*', regex=True, na=False))]
        add1_df = formatting(address1, 1)
        add2_df = formatting(address2, 2)
        self.list = pd.concat([add1_df, add2_df])
        self.list.reset_index(inplace = True, drop=True)       

    def address_standardization(self):
        upperlist = ['Title', 'Dept', 'Address Line 1', 'Address Line 2', 'Postal', 'City', 'Country']
        for item in upperlist:
            self.list[item] = self.list[item].str.upper()
            self.list[item] = self.list[item].str.replace(".", "")
            self.list[item] = self.list[item].str.replace(" - ", "-")
            self.list[item] = self.list[item].str.replace(" -", "-")
            self.list[item] = self.list[item].str.replace("- ", "-")
            self.list[item] = self.list[item].str.replace("-", ", ")
            self.list[item] = self.list[item].str.replace("#", "")
            self.list[item] = self.list[item].str.replace("!", "1")
            self.list[item] = self.list[item].str.replace(r"(^|\b)(C\\O|C\\0|C\/O|C\/0)($|\b)", r'\1\3', regex=True)
            self.list[item] = self.list[item].str.replace(r"(^|\b)(PAV)(\b|$)", r'PAVILLION', regex=True)
            self.list[item] = self.list[item].str.replace(r"(^|\b)(FLR|FL)(\b|$)", r'FLOOR', regex=True)
            self.list[item] = self.list[item].str.replace(r"(^|\b)(RM)(\b|$)", r'ROOM', regex=True)
            self.list[item] = self.list[item].str.replace(r"(^|\b)(CTR)(\b|$)", r'CENTRE', regex=True)
            self.list[item] = self.list[item].str.replace(r"(^|\b)(HLTH)(\b|$)", r'HEALTH', regex=True)
            self.list[item] = self.list[item].str.replace(r"(^|\b)(STE)(\b|$)", r'SUITE', regex=True)
            self.list[item] = self.list[item].str.replace(r"(\b|^)(DEPARTMENT)(\b|$)", r'DEPT', regex=True)
            self.list[item] = self.list[item].str.replace(r"(\b|^)(DRIVE)(\b|$)", r'DR', regex=True)
            self.list[item] = self.list[item].str.replace(r"(^|\b)(BUILDING)(\b|$)", r'BLDG', regex=True)
            self.list[item] = self.list[item].str.replace(r"(\b|^)(ROAD)(\b|$)", r'RD', regex=True)
            self.list[item] = self.list[item].str.replace(r"(\b|^)(CRESCENT)(\b|$)", r'CRES', regex=True)
            self.list[item] = self.list[item].str.replace(r"(\b|^)(AVENUE)(\b|$)", r'AVE', regex=True)
            self.list[item] = self.list[item].str.replace(r"(\b|^)(STREET)(\b|$)", r'ST', regex=True)
            self.list[item] = self.list[item].str.replace(r"(\b|^)(HOSP)(\b|$)", r'HOSPITAL', regex=True)
            self.list[item] = self.list[item].str.replace(r"(\b|^)(1ST)(\b|$)", r'1', regex=True)
            self.list[item] = self.list[item].str.replace(r"([2-9])\s?TH(\b|$)", r'\1', regex=True)
            self.list[item] = self.list[item].str.replace(r"([2-9])\s?ND(\b|$)", r'\1', regex=True)
            self.list[item] = self.list[item].str.replace(r"([2-9])\s?RD(\b|$)", r'\1', regex=True)
            self.list[item] = self.list[item].str.replace(r"(^|\b)(UNIT\s\d+)(\s|$|\b)", r'\1, \2,\3', regex=True)
            self.list[item] = self.list[item].str.replace(r"(^|\b)(SUITE\s\d+)(\s|$|\b)", r'\1, \2,\3', regex=True)
            self.list[item] = self.list[item].str.replace(r"(^|\b)(FLOOR\s\d+)(\s|$|\b)", r'\1, \2,\3', regex=True)
            self.list[item] = self.list[item].str.replace(r"(^|\b)(\d+\sFLOOR)(\s|$|\b)", r'\1, \2,\3', regex=True)
            self.list[item] = self.list[item].str.replace(r"(^|\b)(ROOM\s\d+)(\s|$|\b)", r'\1, \2,\3', regex=True)
            self.list[item] = self.list[item].str.replace(r"(^|\b)(\d+\sROOM)(\s|$|\b)", r'\1, \2,\3', regex=True)
            self.list[item] = self.list[item].str.replace(r"((^|\b)\d+.+(RD|WAY|ST|MALL|BENCH|CRES|DR|AVE|BLDG|HWY|BLVD)\s*([A-Z]|N[A-Z]|E[A-Z]|S[A-Z]|W[A-Z])\s*($|\b))", r', \1, ', regex=True)

    def phone_standardization(self):
        # Phone numbers and faxes adhere to the same format
        phone=['Fax', 'Phone']
        self.list[phone] = self.list[phone].astype(str).applymap(lambda x: re.sub(r'\D', '', x)[-10:])
            
    def postal_standardization(self, upper=False):
        # Makes sure the Postal Codes all adhere to the same format
        col = 'Postal'
        if upper != False:
            col = col.upper()
        self.list[col] = self.list[col].str.replace(' ', '')
        self.list[col] = self.list[col].astype(str).apply(lambda x: (x[:3]+" "+x[3:]) if x != 'NAN' else np.nan)
        self.list[col] = self.list[col].str.replace(r'(^|\b)(NAN|nan)($|\b)', '', regex=True)

    def join_columns(self, column_name, columns):
        self.list[column_name] = self.list[columns].apply(lambda x: ', '.join(x.dropna()), axis=1)

    def po_box_standardization(self):
        # Puts PO BOX in separate place and standardizes
        po_col='PO_BOX'
        addr_col='Address Full'
        self.list[po_col] = self.list[addr_col].str.extract(r'([PO]*\s?BOX\w?.*?\d+|BAG*\s*?\d+)', expand=False)
        self.list[addr_col] = self.list[addr_col].str.replace(r'([PO]*\s?BOX\w?.*?\d+|BAG*\s*?\d+)', '', regex=True)  
        self.list[po_col] = self.list[po_col].str.replace(r'(PO*\s*?BOX.*?)', 'PO BOX', regex=True)   
        self.list[po_col] = self.list[po_col].str.replace(r'(PO BOX )\1+', r'\1', regex=True)   
        self.list[po_col] = self.list[po_col].str.replace(r'(PO BOX,)\1+', r'\1', regex=True)   
        self.list[po_col] = self.list[po_col].str.replace(r'( , )', ' ', regex=True)

    def move_column(self, col, pos):
        col = self.list.pop(col)
        self.list.insert(pos, col.name, col)
    
    def empty_col(self, *args):
        for arg in args:
            self.list[arg] = ''

    def zero_col(self, *args):
        for arg in args:
            self.list[arg] = 0

    def remove_commas(self, col):
        self.list[col] = self.list[col].str.replace(r',+', ',', regex=True)
        self.list[col] = self.list[col].str.replace(r'(^,.*?)', '', regex=True)
        self.list[col] = self.list[col].str.replace(r'( ,)+', '', regex=True)
        self.list[col] = self.list[col].str.replace(r'( , )+', '', regex=True)
        self.list[col] = self.list[col].str.strip()
        self.list[col] = self.list[col].str.rstrip(',')

    def divide(self, col):
        self.list.loc[self.list[col]!='', 'A_SPLIT'] = self.list[col].str.split(',')
        self.list = pd.concat([self.list, self.list[col].str.split(',', expand=True)], axis=1, sort=False)
        mymap = {i: f'A_SPLIT_{i}' for i in range(0, 15)}
        self.list.rename(index=str, columns=mymap, inplace=True)

    def conquer(self, col=None, CPSBC_mod=False):
    # Looks for key words in split columns
        if CPSBC_mod == True:
            for j in range(0, 15):
                try:
                # For Devs, this number may change depending on how many columns addresses get divided into
                    self.list[f'A_SPLIT_{j}'] = self.list[f'A_SPLIT_{j}'].str.strip()
                    self.separate_columns(col_in=f'A_SPLIT_{j}', index_j=j, CPSBC_mod=True)
                except Exception as e:
                    None
        else:
            self.separate_columns(col_in=col)

    def separate_columns(self, col_in, index_j=None, CPSBC_mod=False):
    # Looks for key terms in columns and separates them
        def dept_correction(tempdict):
            for key,value in tempdict.items():
                self.list.loc[value, key] = self.list[col_in].str.strip()
                self.list.loc[value, col_in] = np.nan
                if key == 'DEPT':
                # Separates anything from Hospital (unless its Hospital of _____) and appends it to the Dept column
                    self.list.loc[~(self.list['HOSP'].astype(str).str.contains(r"(.*HOSPITAL\sOF)",regex=True, na=True)), 'HOSP_EXTRACT'] = self.list['HOSP'].astype(str).str.extract(r'.*HOSPITAL\s(.*)', expand=False)
                    self.list.loc[~(self.list['HOSP'].astype(str).str.contains(r"(.*HOSPITAL\sOF)",regex=True, na=False)), 'HOSP'] = self.list['HOSP'].astype(str).str.replace(r'(.*HOSPITAL)\s(.*)', r'\1', regex=True)
                    self.list.loc[(self.list['HOSP'].str.contains(r'(\b|^)nan(\b|$)', regex=True)), 'HOSP'] = ''
                    lists = ['DEPT', 'HOSP_EXTRACT']
                    self.list['DEPT'] = self.list[lists].apply(lambda x: ', '.join(x.dropna()), axis=1)
                    self.remove_commas(col= 'DEPT')
                    self.list.drop(columns=['HOSP_EXTRACT'], inplace=True)

        # Keyword Searches and column separation via dictionaries
        corrections = (self.list[col_in].str.contains(r"((\b|^)(CORRECTION.*?|IMMIGRA.*?|PRETRIAL|CUSTODY|INSTITUTION.*?|DETENTION|HOLDING|HEALING\s*?VILLAGE)(\b|$))",regex=True, na=False))
        first_nations = (self.list[col_in].str.contains(r"((\b|^)(FIRST\s*?NATION.*?|FIRST\s*?PEOPLE.*?|INDIGENOUS|NATIVE|ABORIGINAL|K.LALA\s*?LELUM)(\b|$))",regex=True, na=False))
        admin = (self.list[col_in].str.contains(r"((\b|^)(AIRPORT|CONSULTING|ADMIN.*?|FRASER\sHEALTH\sAUTHORITY|FIRST\sNATION.*?\sHEALTH\sAUTHORITY|CORONER|CPSBC|COLLEGE\sOF\sPHYSICIAN.*?\sAND\sSURGEON.*?|HEALTH\s*?CANADA|VCH|WORKSAFE|WORKSAFEBC|WORKER.*?\s*?COMP.*?|BCAA|VETERAN.*?\s*?AFFAIR.*?|RCMP|AIR\s*?CANADA|QUALITY)(\b|$))",regex=True, na=False))
        # Specialty
        mhsu = (self.list[col_in].str.contains(r"((\b|^)(MENTAL|SUBSTANCE|PSYCH.*?|MHSU|ADDICT.*?|WITHDRAWAL|DEPEND.*?)(\b|$))",regex=True, na=False))
        ltc = (self.list[col_in].str.contains(r"((\b|^)(LODGING|LODGE.*?|MANOR|SENIOR.*?|ALC)(\b|$))",regex=True, na=False))
        derm = (self.list[col_in].str.contains(r"((\b|^)(LASER|SKIN|DERMA.*?)(\b|$))",regex=True, na=False))
        rehab = (self.list[col_in].str.contains(r"((\b|^)(CBI|RECOVERY|MUSCLE|REHAB.*?|SPORT.*?|SPA|OCCUP.*?|PHYSIO.*?)(\b|$))",regex=True, na=False))
        plastics = (self.list[col_in].str.contains(r"((\b|^)(AESTHETIC.*?|COSMETIC|AGING)(\b|$))",regex=True, na=False))
        palliative = (self.list[col_in].str.contains(r"((\b|^)(PALLIATIVE)(\b|$))",regex=True, na=False))
        gyn = (self.list[col_in].str.contains(r"((\b|^)(WOM[AE]N.*?|MENOPAUSE|MATERN.*?|BIRTH.*?|OBSTETRIC.*?|GYNE.*?)(\b|$))",regex=True, na=False))
        pain = (self.list[col_in].str.contains(r"((\b|^)(PAIN)(\b|$))",regex=True, na=False))
        geriatric = (self.list[col_in].str.contains(r"((\b|^)(GERIATRIC)(\b|$))",regex=True, na=False))
        image = (self.list[col_in].str.contains(r"((\b|^)(IMAGING)(\b|$))",regex=True, na=False))
        sex = (self.list[col_in].str.contains(r"((\b|^)(SEXUAL|STI|STD|ELIZABETH\sBAGSHAW)(\b|$))",regex=True, na=False))
        onco = (self.list[col_in].str.contains(r"((\b|^)(CANCER|ONCO.*?)(\b|$))",regex=True, na=False))
        disorder = (self.list[col_in].str.contains(r"((\b|^)(SLEEP|DISORDER)(\b|$))",regex=True, na=False))
        cns = (self.list[col_in].str.contains(r"((\b|^)(SPINE|NEURO.*?|SPINAL)(\b|$))",regex=True, na=False))
        surg = (self.list[col_in].str.contains(r"((\b|^)(SURGERY|SURGICAL)(\b|$))",regex=True, na=False))
        ortho = (self.list[col_in].str.contains(r"((\b|^)(ORTHOP.*?)(\b|$))",regex=True, na=False))
        travel = (self.list[col_in].str.contains(r"((\b|^)(TRAVEL|IMMUNIZ.*?)(\b|$))",regex=True, na=False))
        vascular = (self.list[col_in].str.contains(r"((\b|^)(VEIN|VASCULAR)(\b|$))",regex=True, na=False))
        # End Specialty
        hospital = (self.list[col_in].str.contains(r"((\b|^)(HOSP.*?)(\b|$))",regex=True, na=False))&~(self.list[col_in].str.contains(r"\d",regex=True, na=False))
        dept = (self.list[col_in].str.contains(r"((\b|^)(DEPT.*?|DIV.*?)(\b|$))",regex=True, na=False))
        virtual = (self.list[col_in].str.contains(r"((\b|^)(BABYLON|VIRTUAL|E-*?HEALTH|TELE.*?|I-*?HEALTH.*?)(\b|$))",regex=True, na=False))
        agency = (self.list[col_in].str.contains(r"((\b|^)(AGENCY)(\b|$))",regex=True, na=False))
        university = (self.list[col_in].str.contains(r"((\b|^)(UNIVERSITY|STUDENT|CAMPUS)(\b|$))",regex=True, na=False))&~(self.list[col_in].str.contains(r"\d",regex=True, na=False))
        family_med = (self.list[col_in].str.contains(r"((\b|^)(FAMILY)\b\s*?(MED.*?|CLINIC|CENTRE|CENTER|ASSOCIATE.*?|CARE|PRACTICE|AND\sMATERNITY)(\b|$))",regex=True, na=False))
        clinic = (self.list[col_in].str.contains(r"((\b|^)(CLINIC|ASSOCIATE.*?)(\b|$))",regex=True, na=False))
        centre = (self.list[col_in].str.contains(r"((\b|^)(CENTER|CENTRE|PRACTICE|DOCTOR.*?S)(\b|$))",regex=True, na=False))
        med = (self.list[col_in].str.contains(r"((\b|^)(MED.*?|CARE)(\b|$))",regex=True, na=False))
        society = (self.list[col_in].str.contains(r"((\b|^)(SOCIETY|SERVICE)(\b|$))",regex=True, na=False))
        unit = (self.list[col_in].str.contains(r"((\b|^)(\sUNIT\b$|\sUNIT\b\s\D|PROG.*?|CORP.*?)(\b|$))",regex=True, na=False))
        health = (self.list[col_in].str.contains(r"((\b|^)(HEALTH|WELLNESS)(\b|$))",regex=True, na=False))

        tempdict = {
                    'HOSP':hospital,
                    'FAMILY':family_med,
                    'CORRECTIONS':corrections,
                    'FIRST_NATIONS':first_nations,
                    'SEX':sex, 
                    'GYN':gyn, 
                    'VIRTUAL':virtual,
                    'ADMIN':admin,                
                    'LTC':ltc,
                    'MHSU':mhsu,
                    'DERM':derm, 
                    'REHAB':rehab, 
                    'PLASTICS':plastics, 
                    'PALLIATIVE':palliative, 
                    'PAIN':pain, 
                    'GERIATRIC':geriatric, 
                    'IMAGE':image, 
                    'ONCO':onco, 
                    'DISORDER':disorder, 
                    'CNS':cns, 
                    'SURG':surg, 
                    'ORTHO':ortho, 
                    'TRAVEL':travel,
                    'VASCULAR':vascular,
                    'DEPT':dept, 
                    'AGENCY':agency,
                    'UNI':university,
                    'CLINIC':clinic,
                    'CENTRE':centre,
                    'MED':med,
                    'SOCIETY':society,
                    'UNIT':unit,
                    'HEALTH':health
                    }

        if (index_j!=None):
            mask = ~(self.list[col_in].str.contains(r"(^|\b)\d*.*(RD|WAY|ST|MALL|BENCH|CRES|AVE|HWY|BLVD|BLDG)(\b|$)",regex=True, na=False))
            mask2 = self.list[col_in].str.contains(r"^[a-z\sA-Z\(\)\/\\]+$",regex=True, na=False)
            mask3 = ~(self.list[col_in].str.contains(r"(^|\b)\s*(RD|WAY|ST|MALL|BENCH|CRES|AVE|HWY|BLVD)(\b|$)",regex=True, na=False))
            other = (mask2)&(mask3)
            dept_correction(tempdict=tempdict)
            tempdict[f'OTHER_{index_j}'] = other
            self.list.loc[other, f'OTHER_{index_j}'] = self.list[col_in].str.strip()
            self.list.loc[other, col_in] = np.nan
            self.list.loc[(self.list[col_in].str.contains(r"(^|\b)(PO|BOX)($|\b)",regex=True, na=False)), col_in] = np.nan
            self.list.loc[(self.list[col_in].str.contains(r"^(NORTH|EAST|SOUTH|WEST)$",regex=True, na=False))&(mask), col_in] = np.nan
            self.join_columns('OTHER', ['OTHER', f'OTHER_{index_j}'])
            self.list['OTHER'] = self.list['OTHER'].str.replace(r'(\b|^)(nan|None|NaN)(\b|$)', r'\1\3', regex=True)
            self.list['OTHER'] = self.list['OTHER'].str.strip()
            self.remove_commas(col='OTHER')
            self.join_columns('FORMAL_ADDR', ['FORMAL_ADDR', col_in])
            self.list['FORMAL_ADDR'] = self.list['FORMAL_ADDR'].str.replace(r'(\b|^)(nan|None|NaN)(\b|$)', r'\1\3', regex=True)
            self.list['FORMAL_ADDR'] = self.list['FORMAL_ADDR'].str.strip()
        else:
            dept_correction(tempdict=tempdict)
            self.list.rename(columns={col_in: 'OTHER'}, inplace=True) 

    def bracket_isolation(self):
        # Just in case, finds anything in brackets and extracts it
        self.list['BRACKETS'] = self.list['FORMAL_ADDR'].str.extract(r"(\(.*?\))", expand=False)
        self.list['FORMAL_ADDR'] = self.list['FORMAL_ADDR'].str.replace(r"(\(.*?\))", '', regex=True)

    def geo_prep(self, addr):
        # Prepares an address for geocoding
        self.join_columns(column_name="ADDR_FOR_GEO", columns=addr)
        self.remove_commas(col="ADDR_FOR_GEO")
    
    def geocode(self, geolocator, geocoded):
        self.list['GEO_LOCATION'] = self.list['ADDR_FOR_GEO'].astype(str).apply(lambda doot: geocoded(doot) if doot!= 'nan' else '')
        self.list['GEO_ADDRESS'] = self.list['GEO_LOCATION'].apply(lambda loc: loc.address if loc else "")
        self.list['GEO_RAW'] = self.list['GEO_LOCATION'].apply(lambda loc: loc.raw if loc else "")
        self.list['GEO_GPS'] = self.list['GEO_LOCATION'].apply(lambda loc: loc.point if loc else "")
        self.list['GEO_LATITUDE'] = self.list['GEO_GPS'].apply(lambda loc: loc.latitude if loc else "")
        self.list['GEO_LONGITUDE'] = self.list['GEO_GPS'].apply(lambda loc: loc.longitude if loc else "")
        self.list['FLAG'] = 0
        self.list.loc[(self.list['GEO_ADDRESS'].astype(str).str.contains(r"^[^,]*((,[^,]*){0,1}$)", regex=True, na=True))|~(self.list['GEO_ADDRESS'].astype(str).str.contains(r"\d", regex=True, na=True)), 'FLAG'] = 1
        self.list.loc[(self.list['PO_BOX'].str.contains(r".*", regex=True, na=False))&(self.list['FLAG']==1), ['GEO_LOCATION','GEO_ADDRESS','GEO_RAW','GEO_GPS','GEO_LATITUDE','GEO_LONGITUDE']] = np.nan
        self.list.loc[(self.list['PO_BOX'].str.contains(r".*", regex=True, na=False))&(self.list['FLAG']==1), 'FLAG'] = 0

if __name__ == "__main__":
    writer = pd.ExcelWriter('EXAMPLE_OUTPUT_STEP1.xlsx', engine='openpyxl')
    CPSBC = List('materials\\CPSBC Physician List_2020_09_17.xlsx', 'ResultsGrid_ExportData', CPSBC_mod=True)
    WALKIN = List('materials\\hlbc_walkinclinics.xlsx', 'hlbc_walkinclinics (1)')
    CORR_FAC = List('materials\\corrections_fac.xlsx', 'corr_fac')
    GSR = List('materials\\gsr_assisted_living_feb_25.csv')
    HOSP = List('materials\\hlbc_hospitals.csv')
    UPCC = List('materials\\hlbc_urgentandprimarycarecentres (1).csv')
    QFD = List('materials\\QFD 2020 - public release - 20210105.xlsx', 'QFD 2020')
    CPSBC.drop_and_combine()
    CPSBC.address_standardization()
    CPSBC.phone_standardization() 
    CPSBC.postal_standardization()
    CPSBC.join_columns('Address Full', ['Title', 'Dept', 'Address Line 1', 'Address Line 2'])
    CPSBC.remove_commas('Address Full')
    CPSBC.empty_col('OTHER','FORMAL_ADDR')
    CPSBC.divide('Address Full')
    CPSBC.po_box_standardization()
    CPSBC.move_column('PO_BOX', 4)  
    CPSBC.conquer(CPSBC_mod=True)
    GSR.list['BUSINESS_NAME'] = GSR.list['BUSINESS_NAME'].str.upper()
    HOSP.list['SV_NAME'] = HOSP.list['SV_NAME'].str.upper()
    UPCC.list['SV_NAME'] = UPCC.list['SV_NAME'].str.upper()
    QFD.list['FACILITY_NAME'] = QFD.list['FACILITY_NAME'].str.upper()
    WALKIN.list['RG_NAME'] = WALKIN.list['RG_NAME'].str.upper()
    CORR_FAC.list['RG_NAME'] = CORR_FAC.list['RG_NAME'].str.upper()
    CORR_FAC.list['MAIL_ADDRESS'] = CORR_FAC.list['MAIL_ADDRESS'].str.upper()
    GSR.conquer(col='BUSINESS_NAME')
    HOSP.conquer(col='SV_NAME')
    UPCC.conquer(col='SV_NAME')
    QFD.conquer(col='FACILITY_NAME')
    WALKIN.conquer(col='RG_NAME')
    CORR_FAC.conquer(col='RG_NAME')
    # Misc standardizing
    GSR.empty_col('PO_BOX')
    GSR.list['PROVINCE'] = 'BC'
    HOSP.empty_col('PO_BOX')
    UPCC.empty_col('PO_BOX')
    QFD.empty_col('PO_BOX')
    QFD.list['PROVINCE'] = 'BC'
    WALKIN.empty_col('PO_BOX')
    CORR_FAC.list.rename(columns={'MAIL_ADDRESS': 'PO_BOX'}, inplace = True) 
    for j in range(0, 15):
        try:
            CPSBC.list.drop(columns=[f'A_SPLIT_{j}'], inplace=True)              
        except Exception as e:
            None
        try:
            CPSBC.list.drop(columns=[f'OTHER_{j}'], inplace=True)                
        except Exception as e:
            None
    CPSBC.bracket_isolation()
    CPSBC.geo_prep(addr=['FORMAL_ADDR', 'City', 'Province'])
    WALKIN.geo_prep(addr=['STREET_NUMBER','STREET_NAME','STREET_TYPE','STREET_DIRECTION','CITY','PROVINCE'])
    CORR_FAC.geo_prep(addr=['STREET_NUMBER','CITY','PROVINCE'])
    GSR.geo_prep(addr=['STREET_ADDRESS','CITY','PROVINCE'])
    HOSP.geo_prep(addr=['STREET_NUMBER','STREET_NAME','STREET_TYPE','STREET_DIRECTION','CITY','PROVINCE'])
    UPCC.geo_prep(addr=['STREET_NUMBER','STREET_NAME','STREET_TYPE','STREET_DIRECTION','CITY','PROVINCE'])
    QFD.geo_prep(addr=['STREET_ADDRESS','CITY','PROVINCE'])
    CPSBC.list.drop(columns=['A_SPLIT', 'FORMAL_ADDR'], inplace=True)

    print('Geocoding Start')
    input_val = input('Please Enter Your Company Name (No spaces, abbreviated preferred): ')
    geolocator = DataBC(user_agent=input_val)
    # This delays the rate at which addresses are queried, per DataBC TOS
    geocoded = RateLimiter(geolocator.geocode, min_delay_seconds=1/15)
    print("Geocoding CPSBC List")
    CPSBC.geocode(geolocator, geocoded) 
    print("Geocoding WALKIN List")
    WALKIN.geocode(geolocator, geocoded)
    print("Geocoding Corrections Facilities List")
    CORR_FAC.geocode(geolocator, geocoded)    
    print("Geocoding GSR List")
    GSR.geocode(geolocator, geocoded)   
    print("Geocoding Hospitals List")
    HOSP.geocode(geolocator, geocoded)   
    print("Geocoding UPCC List")
    UPCC.geocode(geolocator, geocoded)   
    print("Geocoding QFD List")
    QFD.geocode(geolocator, geocoded)       

    # Output
    CPSBC.out_sheet(writer, 'CPSBC Modified')
    WALKIN.out_sheet(writer, 'Walk-in Clinic List')
    CORR_FAC.out_sheet(writer, 'Corrections Facilities List')
    GSR.out_sheet(writer, 'GSR List')
    HOSP.out_sheet(writer, 'Hospitals List')
    UPCC.out_sheet(writer, 'UPCC List')
    QFD.out_sheet(writer, 'QFD List')
    writer.save()









    

    
