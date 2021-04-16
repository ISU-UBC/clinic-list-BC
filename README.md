# clinic-list-BC
Clinic list algorithm used to determine a list of primary care clinics in British Columbia, Canada. Version 1 part of a manuscript submitted for publication Spring 2021.

## Authors

### Code:
Cameron Lindsay1 

### Project Authors: 
Ian R. Cooper1, Cameron Lindsay1, Keaton Fraser1, Tiffany T. Hill1, Andrew Siu1, Sarah Fletcher1, Jan Klimas1,2, Michee-Ana Hamilton1, Amanda D. Frazer1, Elka Humphrys1, Kira Koepke1, Lindsay Hedden3,4, Morgan Price1, Rita K. McCracken1

## Author affiliations: 
1.	Innovation Support Unit, Department of Family Practice, University of British Columbia, Vancouver, BC, Canada
2.	British Columbia Centre on Substance Use, Vancouver, BC, Canada
3.	Faculty of Health Sciences, Simon Fraser University, Burnaby, BC, Canada
4.	British Columbia Academic Health Sciences Network, Vancouver, BC, Canada

## Requirements
* Python v3.7.7 or higher
### Libraries
* numpy
* pandas
* geopy
* re

## Materials
### bc_health_region_master_2018.xlsx
The BC Ministry of Health list of regional authorities

### corrections_fac.xlsx
A generated a list of corrections facilities in BC from multiple publicly available sources

### CPSBC Physician List_2020_09_17.xlsx
The College of Physicians and Surgeons of BC list of registered physicians

### gsr_assisted_living_feb_25.csv
A list of long-term care facilities from the BC Office of the Seniors Advocate

### hlbc_hospitals.csv
The BC Ministry of Health list of hospital locations in BC 

### hlbc_urgentandprimarycarecentres (1).csv
The BC Ministry of Health list of urgent primary care clinic locations in BC 

### hlbc_walkinclinics.xlsx
The BC Ministry of Health list of walk-in locations in BC 

### PC_CHSA.csv
The BC Ministry of Health list of regional authorities and corresponding postal codes

### QFD 2020 - public release - 20210105.xlsx
A list of long-term care facilities from the BC Office of the Seniors Advocate

## Code Used (V1)
* camp.py
* filters.py
* final_list.py
* geocode.py
* geocode_added.py
* geocode_flagged.py

## Code Used (V2)
* step1.py
* step3.py

## Steps (V1)
1. Ensure all materials are local (change directory if required), run 'camp.py'
2. This will return 'modified.xlsx'
3. Run 'geocode.py' *NOTE: Geocoding is a time-consuming process and may take upwards of 20-30 minutes as limits are placed on queried addresses.*
4. This will return 'flagged_lists.xlsx'
5. Run 'geocode_added.py'.
6. This will return 'flagged_added.xlsx'
7. Rename 'flagged_lists.xlsx' to 'flagged_lists_manual_edits.xlsx'. Making a copy of the original file is recommended.
8. Manually edit any flagged items to a proper format (column FLAG = 1 in excel file). I would recommend verifying manual edits through [DataBC's geocoding services demo.](https://bcgov.github.io/ols-devkit/ols-demo/index.html)
9. Run 'geocode_flagged.py'.
10. This will return 'working_list.xlsx'
11. Repeat step 8 on the working list if necessary.
12. Run 'filters.py'.
13. This will return 'filtered_list.xlsx'
14. Run 'final_list.py'.
15. This will return 'clinic_list.xlsx'
16. Enjoy!

## Steps (V2)
1. Ensure all materials are in materials folder locally (change directory if required), run 'step1.py'. *NOTE: Geocoding is a time-consuming process and may take upwards of 20-30 minutes as limits are placed on queried addresses.*
2. This will return 'OUTPUT_STEP1.xlsx'
3. Rename 'OUTPUT_STEP1.xlsx' to 'INPUT_STEP2.xlsx'. Making a copy of the original file is recommended.
4. Manually edit any flagged items to a proper format (column FLAG = 1 in excel file). I would recommend verifying manual edits through [DataBC's geocoding services demo.](https://bcgov.github.io/ols-devkit/ols-demo/index.html)
5. Run 'step3.py'.
6. This will return 'FINAL.xlsx'
7. Enjoy!
