import numpy as np
import pandas as pd
import xlwings as xw
import re

def address_cleaning(df, col_in):
    df[col_in] = df[col_in].str.upper()
    df[col_in] = df[col_in].str.replace(".", "")
    df[col_in] = df[col_in].str.replace(" - ", "-")
    df[col_in] = df[col_in].str.replace(" -", "-")
    df[col_in] = df[col_in].str.replace("- ", "-")
    df[col_in] = df[col_in].str.replace("-", ", ")
    df[col_in] = df[col_in].str.replace("#", "")
    df[col_in] = df[col_in].str.replace("!", "1")
    df[col_in] = df[col_in].str.replace(r"(^|\b)(C\\O|C\\0|C\/O|C\/0)($|\b)", r'\1\3', regex=True)
    df[col_in] = df[col_in].str.replace(r"(^|\b)(PAV)(\b|$)", r'PAVILLION', regex=True)
    df[col_in] = df[col_in].str.replace(r"(^|\b)(FLR|FL)(\b|$)", r'FLOOR', regex=True)
    df[col_in] = df[col_in].str.replace(r"(^|\b)(RM)(\b|$)", r'ROOM', regex=True)
    df[col_in] = df[col_in].str.replace(r"(^|\b)(CTR)(\b|$)", r'CENTRE', regex=True)
    df[col_in] = df[col_in].str.replace(r"(^|\b)(HLTH)(\b|$)", r'HEALTH', regex=True)
    df[col_in] = df[col_in].str.replace(r"(^|\b)(STE)(\b|$)", r'SUITE', regex=True)
    df[col_in] = df[col_in].str.replace(r"(\b|^)(DEPARTMENT)(\b|$)", r'DEPT', regex=True)
    df[col_in] = df[col_in].str.replace(r"(\b|^)(DRIVE)(\b|$)", r'DR', regex=True)
    df[col_in] = df[col_in].str.replace(r"(^|\b)(BUILDING)(\b|$)", r'BLDG', regex=True)
    df[col_in] = df[col_in].str.replace(r"(\b|^)(ROAD)(\b|$)", r'RD', regex=True)
    df[col_in] = df[col_in].str.replace(r"(\b|^)(CRESCENT)(\b|$)", r'CRES', regex=True)
    df[col_in] = df[col_in].str.replace(r"(\b|^)(AVENUE)(\b|$)", r'AVE', regex=True)
    df[col_in] = df[col_in].str.replace(r"(\b|^)(STREET)(\b|$)", r'ST', regex=True)
    df[col_in] = df[col_in].str.replace(r"(\b|^)(HOSP)(\b|$)", r'HOSPITAL', regex=True)
    df[col_in] = df[col_in].str.replace(r"(\b|^)(1ST)(\b|$)", r'1', regex=True)
    df[col_in] = df[col_in].str.replace(r"([2-9])\s?TH(\b|$)", r'\1', regex=True)
    df[col_in] = df[col_in].str.replace(r"([2-9])\s?ND(\b|$)", r'\1', regex=True)
    df[col_in] = df[col_in].str.replace(r"([2-9])\s?RD(\b|$)", r'\1', regex=True)
    df[col_in] = df[col_in].str.replace(r"(^|\b)(UNIT\s\d+)(\s|$|\b)", r'\1, \2,\3', regex=True)
    df[col_in] = df[col_in].str.replace(r"(^|\b)(SUITE\s\d+)(\s|$|\b)", r'\1, \2,\3', regex=True)
    df[col_in] = df[col_in].str.replace(r"(^|\b)(FLOOR\s\d+)(\s|$|\b)", r'\1, \2,\3', regex=True)
    df[col_in] = df[col_in].str.replace(r"(^|\b)(\d+\sFLOOR)(\s|$|\b)", r'\1, \2,\3', regex=True)
    df[col_in] = df[col_in].str.replace(r"(^|\b)(ROOM\s\d+)(\s|$|\b)", r'\1, \2,\3', regex=True)
    df[col_in] = df[col_in].str.replace(r"(^|\b)(\d+\sROOM)(\s|$|\b)", r'\1, \2,\3', regex=True)
    df[col_in] = df[col_in].str.replace(r"((^|\b)\d+.+(RD|WAY|ST|MALL|BENCH|CRES|DR|AVE|BLDG|HWY|BLVD)\s*([A-Z]|N[A-Z]|E[A-Z]|S[A-Z]|W[A-Z])\s*($|\b))", r', \1, ', regex=True)


# Looks for key terms in columns and separates them
def separate_columns(df, col_in, index_i=False, index_j=False):
    corrections = (df[col_in].str.contains(r"((\b|^)(CORRECTION.*?|IMMIGRA.*?|PRETRIAL|CUSTODY|INSTITUTION.*?|DETENTION|HOLDING|HEALING\s*?VILLAGE)(\b|$))",regex=True, na=False))
    first_nations = (df[col_in].str.contains(r"((\b|^)(FIRST\s*?NATION.*?|FIRST\s*?PEOPLE.*?|INDIGENOUS|NATIVE|ABORIGINAL|K.LALA\s*?LELUM)(\b|$))",regex=True, na=False))
    admin = (df[col_in].str.contains(r"((\b|^)(AIRPORT|CONSULTING|ADMIN.*?|FRASER\sHEALTH\sAUTHORITY|FIRST\sNATION.*?\sHEALTH\sAUTHORITY|CORONER|CPSBC|COLLEGE\sOF\sPHYSICIAN.*?\sAND\sSURGEON.*?|HEALTH\s*?CANADA|VCH|WORKSAFE|WORKSAFEBC|WORKER.*?\s*?COMP.*?|BCAA|VETERAN.*?\s*?AFFAIR.*?|RCMP|AIR\s*?CANADA|QUALITY)(\b|$))",regex=True, na=False))
    # Specialty
    mhsu = (df[col_in].str.contains(r"((\b|^)(MENTAL|SUBSTANCE|PSYCH.*?|MHSU|ADDICT.*?|WITHDRAWAL|DEPEND.*?)(\b|$))",regex=True, na=False))
    ltc = (df[col_in].str.contains(r"((\b|^)(LODGING|LODGE.*?|MANOR|SENIOR.*?|ALC)(\b|$))",regex=True, na=False))
    derm = (df[col_in].str.contains(r"((\b|^)(LASER|SKIN|DERMA.*?)(\b|$))",regex=True, na=False))
    rehab = (df[col_in].str.contains(r"((\b|^)(CBI|RECOVERY|MUSCLE|REHAB.*?|SPORT.*?|SPA|OCCUP.*?|PHYSIO.*?)(\b|$))",regex=True, na=False))
    plastics = (df[col_in].str.contains(r"((\b|^)(AESTHETIC.*?|COSMETIC|AGING)(\b|$))",regex=True, na=False))
    palliative = (df[col_in].str.contains(r"((\b|^)(PALLIATIVE)(\b|$))",regex=True, na=False))
    gyn = (df[col_in].str.contains(r"((\b|^)(WOM[AE]N.*?|MENOPAUSE|MATERN.*?|BIRTH.*?|OBSTETRIC.*?|GYNE.*?)(\b|$))",regex=True, na=False))
    pain = (df[col_in].str.contains(r"((\b|^)(PAIN)(\b|$))",regex=True, na=False))
    geriatric = (df[col_in].str.contains(r"((\b|^)(GERIATRIC)(\b|$))",regex=True, na=False))
    image = (df[col_in].str.contains(r"((\b|^)(IMAGING)(\b|$))",regex=True, na=False))
    sex = (df[col_in].str.contains(r"((\b|^)(SEXUAL|STI|STD|ELIZABETH\sBAGSHAW)(\b|$))",regex=True, na=False))
    onco = (df[col_in].str.contains(r"((\b|^)(CANCER|ONCO.*?)(\b|$))",regex=True, na=False))
    disorder = (df[col_in].str.contains(r"((\b|^)(SLEEP|DISORDER)(\b|$))",regex=True, na=False))
    cns = (df[col_in].str.contains(r"((\b|^)(SPINE|NEURO.*?|SPINAL)(\b|$))",regex=True, na=False))
    surg = (df[col_in].str.contains(r"((\b|^)(SURGERY|SURGICAL)(\b|$))",regex=True, na=False))
    ortho = (df[col_in].str.contains(r"((\b|^)(ORTHOP.*?)(\b|$))",regex=True, na=False))
    travel = (df[col_in].str.contains(r"((\b|^)(TRAVEL|IMMUNIZ.*?)(\b|$))",regex=True, na=False))
    vascular = (df[col_in].str.contains(r"((\b|^)(VEIN|VASCULAR)(\b|$))",regex=True, na=False))
    # End Specialty
    hospital = (df[col_in].str.contains(r"((\b|^)(HOSP.*?)(\b|$))",regex=True, na=False))&~(df[col_in].str.contains(r"\d",regex=True, na=False))
    dept = (df[col_in].str.contains(r"((\b|^)(DEPT.*?|DIV.*?)(\b|$))",regex=True, na=False))
    virtual = (df[col_in].str.contains(r"((\b|^)(BABYLON|VIRTUAL|E-*?HEALTH|TELE.*?|I-*?HEALTH.*?)(\b|$))",regex=True, na=False))
    agency = (df[col_in].str.contains(r"((\b|^)(AGENCY)(\b|$))",regex=True, na=False))
    university = (df[col_in].str.contains(r"((\b|^)(UNIVERSITY|STUDENT|CAMPUS)(\b|$))",regex=True, na=False))&~(df[col_in].str.contains(r"\d",regex=True, na=False))
    family_med = (df[col_in].str.contains(r"((\b|^)(FAMILY)\b\s*?(AND\s*?MATERNITY|MED.*?|CLINIC|CENTRE|CENTER|ASSOCIATE.*?|CARE|PRACTICE)(\b|$))",regex=True, na=False))
    clinic = (df[col_in].str.contains(r"((\b|^)(CLINIC.*?|ASSOCIATE.*?)(\b|$))",regex=True, na=False))
    centre = (df[col_in].str.contains(r"((\b|^)(CENTER|CENTRE|PRACTICE|DOCTOR.*?S)(\b|$))",regex=True, na=False))
    med = (df[col_in].str.contains(r"((\b|^)(MED.*?|CARE)(\b|$))",regex=True, na=False))
    society = (df[col_in].str.contains(r"((\b|^)(SOCIETY|SERVICE)(\b|$))",regex=True, na=False))
    unit = (df[col_in].str.contains(r"((\b|^)(\sUNIT\b$|\sUNIT\b\s\D|PROG.*?|CORP.*?)(\b|$))",regex=True, na=False))
    health = (df[col_in].str.contains(r"((\b|^)(HEALTH|WELLNESS)(\b|$))",regex=True, na=False))

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

    if (index_i!=False):
        mask = ~(df[col_in].str.contains(r"(^|\b)\d*.*(RD|WAY|ST|MALL|BENCH|CRES|AVE|HWY|BLVD|BLDG)(\b|$)",regex=True, na=False))
        mask2 = df[col_in].str.contains(r"^[a-z\sA-Z\(\)\/\\]+$",regex=True, na=False)
        mask3 = df[col_in].str.contains(r"(^|\b)\s*(RD|WAY|ST|MALL|BENCH|CRES|AVE|HWY|BLVD)(\b|$)",regex=True, na=False)
        other = (mask2)&~(mask3)

        for key,value in tempdict.items():
            key = key+"_{}".format(index_i)
            df.loc[value, key] = df[col_in].str.strip()
            df.loc[value, col_in] = np.nan
            if key == 'DEPT'+"_{}".format(index_i):
            # Separates anything from Hospital (unless its Hospital of _____) and appends it to the Dept column
                df.loc[~(df[f'HOSP_{index_i}'].str.contains(r"(.*HOSPITAL\sOF)",regex=True, na=True)), f'HOSP_EXTRACT_{index_i}'] = df[f'HOSP_{index_i}'].str.extract(r'.*HOSPITAL\s(.*)', expand=False)
                df.loc[~(df[f'HOSP_{index_i}'].str.contains(r"(.*HOSPITAL\sOF)",regex=True, na=False)), f'HOSP_{index_i}'] = df[f'HOSP_{index_i}'].str.replace(r'(.*HOSPITAL)\s(.*)', r'\1', regex=True)
                lists = [f'DEPT_{index_i}', f'HOSP_EXTRACT_{index_i}']
                df[f'DEPT_{index_i}'] = df[lists].apply(lambda x: ', '.join(x.dropna()), axis=1)
                ditch_commas(df, f'DEPT_{index_i}')
                df.drop(columns=[f'HOSP_EXTRACT_{index_i}'], inplace=True)
        
        tempdict[f'OTHER_{index_i}_{index_j}'] = other
        df.loc[other, f'OTHER_{index_i}_{index_j}'] = df[col_in].str.strip()
        df.loc[other, col_in] = np.nan

        df.loc[(df[col_in].str.contains(r"(^|\b)(PO|BOX)($|\b)",regex=True, na=False)), col_in] = np.nan
        df.loc[(df[col_in].str.contains(r"^(NORTH|EAST|SOUTH|WEST)$",regex=True, na=False))&(mask), col_in] = np.nan
        df[f'OTHER_{index_i}'] = df[f'OTHER_{index_i}'].astype(str) + ', ' + df[f'OTHER_{index_i}_{index_j}'].astype(str) 
        df[f'OTHER_{index_i}'] = df[f'OTHER_{index_i}'].str.replace(r'(\b|^)(nan|None|NaN)(\b|$)', r'\1\3', regex=True)
        df[f'OTHER_{index_i}'] = df[f'OTHER_{index_i}'].str.strip()
        ditch_commas(df, f'OTHER_{index_i}')
        df[f'FORMAL_ADDR_{index_i}'] = df[f'FORMAL_ADDR_{index_i}'].astype(str) + ', ' + df[col_in].astype(str) 
        df[f'FORMAL_ADDR_{index_i}'] = df[f'FORMAL_ADDR_{index_i}'].str.replace(r'(\b|^)(nan|None|NaN)(\b|$)', r'\1\3', regex=True)
        df[f'FORMAL_ADDR_{index_i}'] = df[f'FORMAL_ADDR_{index_i}'].str.strip()

    else:

        for key,value in tempdict.items():
            df.loc[value, key] = df[col_in].str.strip()
            df.loc[value, col_in] = np.nan
            if key == 'DEPT':
            # Separates anything from Hospital (unless its Hospital of _____) and appends it to the Dept column
                df.loc[~(df['HOSP'].astype(str).str.contains(r"(.*HOSPITAL\sOF)",regex=True, na=True)), 'HOSP_EXTRACT'] = df['HOSP'].astype(str).str.extract(r'.*HOSPITAL\s(.*)', expand=False)
                df.loc[~(df['HOSP'].astype(str).str.contains(r"(.*HOSPITAL\sOF)",regex=True, na=False)), 'HOSP'] = df['HOSP'].astype(str).str.replace(r'(.*HOSPITAL)\s(.*)', r'\1', regex=True)
                lists = ['DEPT', 'HOSP_EXTRACT']
                df['DEPT'] = df[lists].apply(lambda x: ', '.join(x.dropna()), axis=1)
                ditch_commas(df, 'DEPT')
                df.drop(columns=['HOSP_EXTRACT'], inplace=True)

        df.rename(columns={col_in: 'OTHER'}, inplace=True)

# Moves a specified column to a new location within dataframe
def move_column(df, col, pos):
    col = df.pop(col)
    df.insert(pos, col.name, col)

# Removes excess columns
def ditch_commas(df, col):
    df[col] = df[col].str.replace(r',+', ',', regex=True)
    df[col] = df[col].str.replace(r'(^,.*?)', '', regex=True)
    df[col] = df[col].str.replace(r'( ,)+', '', regex=True)
    df[col] = df[col].str.replace(r'( , )+', '', regex=True)
    df[col] = df[col].str.strip()
    df[col] = df[col].str.rstrip(',')

if __name__ == "__main__":
    # This lets me open the file and enter in the password
    # workbook = xw.Book('CPSBC Physician List_2020_09_17.xlsx')
    # sheet = workbook.sheets['ResultsGrid_ExportData']
    # full_list = sheet['A:AJ'].options(pd.DataFrame, index=False, header = True).value.dropna(how = 'all')
    full_list = pd.read_excel('CPSBC Physician List_2020_09_17.xlsx', sheet_name='ResultsGrid_ExportData').fillna('')
    hlbc_list = pd.read_excel('hlbc_walkinclinics.xlsx', sheet_name='hlbc_walkinclinics (1)')
    corr_fac_list = pd.read_excel('corrections_fac.xlsx', sheet_name='corr_fac')
    writer = pd.ExcelWriter('modified.xlsx', engine='openpyxl')

    # Applying 'cleaning' of addresses:
    # Making Everything Uppercase 
    for i in range(1, 3):
        upperlist = [f'Title {i}', f'Dept {i}', f'Address {i} Line 1', f'Address {i} Line 2', f'Postal {i}', f'City {i}', f'Country {i}']
        for item in upperlist:
            address_cleaning(full_list, item)
 
    # Standardizing the phone numbers
    for i in range(1, 3):
        phone = [f'Fax {i}', f'Phone {i}']
        full_list[phone] = full_list[phone].astype(str).applymap(lambda x: re.sub(r'\D', '', x)[-10:])

    # Makes sure the Postal Codes all adhere to the same format
        full_list[f'Postal {i}'] = full_list[f'Postal {i}'].str.replace(' ', '')
        full_list[f'Postal {i}'] = full_list[f'Postal {i}'].astype(str).apply(lambda x: (x[:3]+" "+x[3:]) if x != None else None)  
 

    # This merges the addresses together   
    for i in range(1, 3):
        address = [f'Title {i}', f'Dept {i}', f'Address {i} Line 1', f'Address {i} Line 2']
        full_list[f'Address {i} Full'] = full_list[address].apply(lambda x: ', '.join(x.dropna()), axis=1)


    # Puts PO BOX in separate place and standardizes
    for i in range(1, 3):
        full_list[f'PO_BOX_ADDR {i}'] = full_list[f'Address {i} Full'].str.extract(r'([PO]*\s?BOX\w?.*?\d+|BAG*\s*?\d+)', expand=False)
        full_list[f'Address {i} Full'] = full_list[f'Address {i} Full'].str.replace(r'([PO]*\s?BOX\w?.*?\d+|BAG*\s*?\d+)', '', regex=True)  
        full_list[f'PO_BOX_ADDR {i}'] = full_list[f'PO_BOX_ADDR {i}'].str.replace(r'(PO*\s*?BOX.*?)', 'PO BOX', regex=True)   
        full_list[f'PO_BOX_ADDR {i}'] = full_list[f'PO_BOX_ADDR {i}'].str.replace(r'(PO BOX )\1+', r'\1', regex=True)   
        full_list[f'PO_BOX_ADDR {i}'] = full_list[f'PO_BOX_ADDR {i}'].str.replace(r'(PO BOX,)\1+', r'\1', regex=True)   
        full_list[f'PO_BOX_ADDR {i}'] = full_list[f'PO_BOX_ADDR {i}'].str.replace(r'( , )', ' ', regex=True)

    move_column(full_list, 'PO_BOX_ADDR 1', 16)  
    move_column(full_list, 'PO_BOX_ADDR 2', 27)  

    # Removes excess commas
    for i in range(1, 3):
        ditch_commas(full_list, f"Address {i} Full")

    # Splits the 'full' address by commas --> this'll make searches for specific terms easier 
        full_list.loc[full_list[f"Address {i} Full"]!='', f'A{i}_SPLIT'] = full_list[f"Address {i} Full"].str.split(',')
        full_list = pd.concat([full_list, full_list[f"Address {i} Full"].str.split(',', expand=True)], axis=1, sort=False)
        mymap = {j: f'A{i}_SPLIT_{j}' for j in range(0, 15)}
        full_list.rename(index=str, columns=mymap, inplace=True)

    # Looks for key words in split columns
    for i in range(1, 3):
        full_list[f'OTHER_{i}'] = np.nan
        full_list[f'FORMAL_ADDR_{i}'] = np.nan
        try:
            for j in range(0, 15):
                full_list[f'A{i}_SPLIT_{j}'] = full_list[f'A{i}_SPLIT_{j}'].str.strip()
                separate_columns(full_list, f'A{i}_SPLIT_{j}', index_i=i,index_j=j)
    
        except Exception as e:
            # print("in separate columns: ", e)
            None
    
    # And in the other lists
    hlbc_list['RG_NAME'] = hlbc_list['RG_NAME'].str.upper()
    corr_fac_list['RG_NAME'] = corr_fac_list['RG_NAME'].str.upper()
    corr_fac_list['MAIL_ADDRESS'] = corr_fac_list['MAIL_ADDRESS'].str.upper()
    separate_columns(hlbc_list, 'RG_NAME')
    separate_columns(corr_fac_list, 'RG_NAME')

    # Dropping the split columns 
    for i in range(1, 3):
        try:
            for j in range(0, 15):
                full_list.drop(columns=[f'A{i}_SPLIT_{j}'], inplace=True)  
                full_list.drop(columns=[f'OTHER_{i}_{j}'], inplace=True)               
        except Exception as e:
            # print("in dropping columns: ", e)
            None
    
    # Just in case, finds anything in brackets and extracts it
    for i in range(1, 3):
        full_list[f'BRACKETS_{i}'] = full_list[f'FORMAL_ADDR_{i}'].str.extract(r"(\(.*?\))", expand=False)
        full_list[f'FORMAL_ADDR_{i}'] = full_list[f'FORMAL_ADDR_{i}'].str.replace(r"(\(.*?\))", '', regex=True)

    # Joins columns for geocoding
        address = [f'FORMAL_ADDR_{i}', f'City {i}', f'Province {i}']
        full_list[f"ADDR_FOR_GEO_{i}"] = full_list[address].apply(lambda x: ', '.join(x.dropna()), axis=1)

    # Removes excess commas for geocoding
        ditch_commas(full_list, f"ADDR_FOR_GEO_{i}")

    # Makes sure geocoded values are BC Specialists
        full_list.loc[(full_list[f'Province {i}']!='BC')|~(full_list['Specialties & Certifications'].str.contains(r'^Fam.*', regex=True, na=False)), f"ADDR_FOR_GEO_{i}"] = ''

    # Dropping no longer used columns
        full_list.drop(columns=[f'A{i}_SPLIT'], inplace=True)
        full_list.drop(columns=[f'FORMAL_ADDR_{i}'], inplace=True)

    # Applying years practice, 2 years added for fam med training???
    full_list['years_in_prac'] = (2020 - full_list['Degree Year'].astype(int) + 2)

    # Applying if they trained in Canada, IMG = international medical grad
    full_list['grad_IMG'] = ~full_list['University'].str.contains('Canada', regex = False)
    full_list['grad_IMG'].replace(True, 1, inplace = True)
    full_list['grad_IMG'].replace(False, 0, inplace = True)  

    # Exports Excel
    # full_list.to_excel('modified.xlsx', index=False)
    full_list.to_excel(writer, sheet_name = 'modified', index=False)
    hlbc_list.to_excel(writer, sheet_name = 'HLBC Clinic List', index=False)
    corr_fac_list.to_excel(writer, sheet_name = 'Corrections Facilities', index=False)
    writer.save()









    

    
