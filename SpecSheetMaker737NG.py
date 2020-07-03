import cx_Oracle
import xlsxwriter
from datetime import datetime, timedelta
import os
import re
import pandas as pd

#Dependencies: Operator, Config Matrix Values, Liveries, Maintenix Data, Reconfigurations of seats, ovens, galleys, A-CKS, C-CKS intervals,
#Connection to intranet (e.g. via VPN)

#Defining valid aircraft list
aircraft_list = ['HP-1371CMP','HP-1372CMP','HP-1373CMP','HP-1374CMP','HP-1375CMP',
'HP-1376CMP','HP-1377CMP','HP-1378CMP','HP-1379CMP','HP-1380CMP','HP-1520CMP','HP-1521CMP','HP-1524CMP',
'HP-1525CMP','HP-1527CMP','HP-1528CMP','HP-1530CMP','HP-1531CMP','HP-1522CMP','HP-1523CMP','HP-1526CMP',
'HP-1532CMP','HP-1533CMP','HP-1534CMP','HP-1535CMP','HP-1536CMP','HP-1537CMP','HP-1538CMP',
'HP-1539CMP','HP-1711CMP','HP-1712CMP','HP-1713CMP','HP-1714CMP','HP-1715CMP','HP-1716CMP','HP-1717CMP',
'HP-1718CMP','HP-1719CMP','HP-1720CMP','HP-1721CMP','HP-1722CMP','HP-1723CMP','HP-1724CMP','HP-1725CMP',
'HP-1726CMP','HP-1727CMP','HP-1728CMP','HP-1729CMP','HP-1730CMP','HP-1821CMP','HP-1822CMP','HP-1823CMP',
'HP-1824CMP','HP-1825CMP','HP-1826CMP','HP-1827CMP','HP-1828CMP','HP-1829CMP','HP-1830CMP','HP-1831CMP',
'HP-1832CMP','HP-1833CMP','HP-1834CMP','HP-1835CMP','HP-1836CMP','HP-1837CMP','HP-1838CMP','HP-1839CMP',
'HP-1840CMP','HP-1841CMP','HP-1842CMP','HP-1843CMP','HP-1844CMP','HP-1845CMP','HP-1846CMP','HP-1847CMP',
'HP-1848CMP','HP-1849CMP','HP-1850CMP','HP-1851CMP','HP-1852CMP','HP-1853CMP','HP-1854CMP','HP-1855CMP',
'HP-1856CMP','HP-1857CMP']

#Requesting and validating correct aircraft input
while True:
    aircraft = input('Type a valid 737NG aircraft registration in format HP-XXXXCMP: ')
    if aircraft_list.count(aircraft) == 1:
        print(f'Generating Aircraft Specification Sheet for {aircraft}...')
        break
    
aircraftnum = str(''.join(list(filter(str.isdigit, aircraft)))) #String manipulation for Config Matrix file

#Date of report generation
today = datetime.today().strftime('%d-%b-%y')

#Setting paths for Config Matrix, Engine Removals, and APU removals spreadsheets (dataframes)
##path_cm = r'/Users/Giancarlo/OneDrive - Compañía Panameña de Aviación, S.A/My Documents/Pandas Test Config Matrix/Config Matrix NG.xlsx'
path_cm = r'\\ATO-DFS-03\WorkfilesTUM\MANTO\DISCOM\ENGINEER\REFERENCE\SpecSheetMaker737NG\Config Matrix NG.xlsx'
##path_er = r'/Users/Giancarlo/AppData/Local/Programs/Python/Python38-32/Engine Removals.xlsx'
path_er = r'\\ATO-DFS-03\WorkfilesTUM\MANTO\DISCOM\ENGINEER\REFERENCE\SpecSheetMaker737NG\Engine Removals.xlsx'
##path_ar = r'/Users/Giancarlo/AppData/Local/Programs/Python/Python38-32/APU Removals.xlsx'
path_ar = r'\\ATO-DFS-03\WorkfilesTUM\MANTO\DISCOM\ENGINEER\REFERENCE\SpecSheetMaker737NG\APU Removals.xlsx'

#Reading from Config Matrix Dataframe
df = pd.read_excel(path_cm,index_col ='Title')

mtw = df.loc[['Maximum Taxi Weight (lbs)'], aircraftnum][0]
mtow = df.loc[['Maximum Take-Off Weight (lbs)'], aircraftnum][0]
mlw = df.loc[['Maximum Landing Weight (lbs)'], aircraftnum][0]
mzfw = df.loc[['Maximum Zero Fuel Weight (lbs)'], aircraftnum][0]
noise_cat = df.loc[['Noise Category'], aircraftnum][0]
cat_status = df.loc[['Landing category approval'], aircraftnum][0]
sfp = df.loc[['Short Field Performance'], aircraftnum][0]
satcom = df.loc[['Iridium'], aircraftnum][0]
hfdle = df.loc[['HFDL enabled'], aircraftnum][0]
hfdlos = df.loc[['HFDL Override switch'], aircraftnum][0]
batt = df.loc[['Battery '], aircraftnum][0]
dfdau_sw = df.loc[['DFDAU Mandatory Software'], aircraftnum][0]
rips= df.loc[['Recorder Independent Power Supply (RIPS)'], aircraftnum][0]
tcas_sw = df.loc[['TCAS Software'], aircraftnum][0]
adf = df.loc[['ADF Receiver'], aircraftnum][0]
ife = df.loc[['In-Flight Entertainment  (IFE)'], aircraftnum][0]
bsi = df.loc[['BSI Interior'], aircraftnum][0]
pax_config = df.loc[['Passengers BC/TC'], aircraftnum][0]
seats_mfg = df.loc[['Seats Manufacturer BC/TC'], aircraftnum][0]
seats_pitch = df.loc[['Seats pitch BC/TC'], aircraftnum][0]
seats_recline = df.loc[['Seats recline BC/TC'], aircraftnum][0]
g1 = df.loc[['Galley G1'], aircraftnum][0]
g2 = df.loc[['Galley G2'], aircraftnum][0]
g4b = df.loc[['Galley G4B'], aircraftnum][0]
g7 = df.loc[['Galley G7'], aircraftnum][0]
lav_config = df.loc[['Lavatory configuration'], aircraftnum][0]
lav_mfg = df.loc[['Lavatory manufacturer'], aircraftnum][0]
oven = df.loc[['Ovens'], aircraftnum][0]
pattern_oven = re.compile(r'PN\s\S*') #Aqui empieza la limpieza del campo de oven
matches = pattern_oven.findall(oven)
oven_g2 = matches[0]
oven_g4 = matches[1]
pattern_mfg = re.compile(r'^[\D]+?\s')
matches = pattern_mfg.findall(oven)
oven_g2 = '2 X ' + matches[0] + oven_g2
oven_g4 = '3 X ' +matches[0] + oven_g4 #Aqui termina
slides = df.loc[['Escape slides'], aircraftnum][0]
elt_fixed = df.loc[['Fixed Automatic ELT'], aircraftnum][0]
elt_port = df.loc[['Portable ELT'], aircraftnum][0]
aft = df.loc[['Auxiliary Fuel Tanks'], aircraftnum][0]
brakes_type = df.loc[['Brakes material'], aircraftnum][0]
brakes_mfgpn = df.loc[['Brakes manufacturer'], aircraftnum][0]
wheels_mfgpn = df.loc[['Main Wheels'], aircraftnum][0]
lg_switch = df.loc[['Swich for dispatch w/ LG down'], aircraftnum][0]
oxy_gen = df.loc[['22-minutes Chemical Oxygen Generators'], aircraftnum][0]
obs_mask = df.loc[['First Observer Full-Face Mask'], aircraftnum][0]
water_tank = df.loc[['Potable Water Capacity'], aircraftnum][0]
ngs = df.loc[['Nitrogen Generation System'], aircraftnum][0]
fd_door = df.loc[['Enhance security cockpit door'], aircraftnum][0] + ',' + df.loc[['Cockpit Door OEM'], aircraftnum][0]
winglets = df.loc[['Winglets'], aircraftnum][0]
dfdau_sw=df.loc[['DFDAU Mandatory Software'],aircraftnum][0]
rips=df.loc[['Recorder Independent Power Supply (RIPS)'],aircraftnum][0]
ife=df.loc[['In-Flight Entertainment  (IFE)'],aircraftnum][0]
tcas_sw=df.loc[['TCAS Software'],aircraftnum][0]


#Oracle SQL Queries Definitions

query_ac_info = f'''SELECT INV_AC_REG.AC_REG_CD, INV_INV.MANUFACT_DT, INV_AC_REG.LINE_NO_OEM, INV_AC_REG.VAR_NO_OEM, INV_INV.SERIAL_NO_OEM,
EQP_PART_NO.PART_NO_OEM AS AC_MODEL

FROM INV_AC_REG

INNER JOIN INV_INV ON
INV_AC_REG.INV_NO_ID = INV_INV.INV_NO_ID

INNER JOIN EQP_PART_NO ON
INV_INV.PART_NO_ID = EQP_PART_NO.PART_NO_ID

WHERE AC_REG_CD = '{aircraft}'
'''


query_tsn_fc = f'''SELECT 
INV_CURR_USAGE.TSN_QT

FROM INV_AC_REG 

INNER JOIN INV_CURR_USAGE  ON 
INV_AC_REG.INV_NO_ID = INV_CURR_USAGE.INV_NO_ID 

WHERE AC_REG_CD = '{aircraft}' AND
INV_CURR_USAGE.DATA_TYPE_ID = 10
'''

query_tsn_fh = f'''SELECT 
INV_CURR_USAGE.TSN_QT

FROM INV_AC_REG 

INNER JOIN INV_CURR_USAGE  ON 
INV_AC_REG.INV_NO_ID = INV_CURR_USAGE.INV_NO_ID 

WHERE AC_REG_CD = '{aircraft}' AND
INV_CURR_USAGE.DATA_TYPE_ID = 1

'''

query_nextcck = f'''SELECT EVT_SCHED_DEAD.SCHED_DEAD_DT
FROM TASK_TASK

INNER JOIN SCHED_STASK ON
TASK_TASK.TASK_DB_ID = SCHED_STASK.TASK_DB_ID AND
TASK_TASK.TASK_ID = SCHED_STASK.TASK_ID

INNER JOIN INV_AC_REG ON
SCHED_STASK.MAIN_INV_NO_ID = INV_AC_REG.INV_NO_ID AND
SCHED_STASK.MAIN_INV_NO_DB_ID = INV_AC_REG.INV_NO_DB_ID

INNER JOIN EVT_SCHED_DEAD ON 
sched_stask.sched_db_id = EVT_SCHED_DEAD.event_db_id  AND
sched_stask.sched_id = EVT_SCHED_DEAD.event_id 

INNER JOIN EVT_EVENT ON 
EVT_SCHED_DEAD.EVENT_ID = EVT_EVENT.EVENT_ID AND
EVT_SCHED_DEAD.EVENT_DB_ID = EVT_EVENT.EVENT_DB_ID

WHERE TASK_TASK.TASK_CD = 'C-CK-1 - 737-NG' AND
EVT_EVENT.EVENT_STATUS_CD = 'ACTV' AND
EVT_SCHED_DEAD.USAGE_REM_QT <= 1095 AND
INV_AC_REG.AC_REG_CD = '{aircraft}'
'''

query_lastcck = f'''
SELECT
EVT_SCHED_DEAD.SCHED_DEAD_DT


FROM TASK_TASK

INNER JOIN SCHED_STASK ON
TASK_TASK.TASK_DB_ID = SCHED_STASK.TASK_DB_ID AND
TASK_TASK.TASK_ID = SCHED_STASK.TASK_ID

INNER JOIN INV_AC_REG ON
SCHED_STASK.MAIN_INV_NO_ID = INV_AC_REG.INV_NO_ID AND
SCHED_STASK.MAIN_INV_NO_DB_ID = INV_AC_REG.INV_NO_DB_ID

INNER JOIN EVT_SCHED_DEAD ON 
sched_stask.sched_db_id = EVT_SCHED_DEAD.event_db_id  AND
sched_stask.sched_id = EVT_SCHED_DEAD.event_id 

INNER JOIN EVT_EVENT ON 
EVT_SCHED_DEAD.EVENT_ID = EVT_EVENT.EVENT_ID AND
EVT_SCHED_DEAD.EVENT_DB_ID = EVT_EVENT.EVENT_DB_ID

WHERE TASK_TASK.TASK_CD = 'C-CK-1 - 737-NG' AND
EVT_EVENT.EVENT_STATUS_CD = 'COMPLETE' AND
INV_AC_REG.AC_REG_CD = '{aircraft}' AND
to_date(CURRENT_DATE) - to_date(EVT_SCHED_DEAD.SCHED_DEAD_DT) <=1095 AND
EVT_SCHED_DEAD.DATA_TYPE_ID = 21

'''

query_nlg_fh = f'''

SELECT 
EQP_PART_NO.PART_NO_OEM,
II.SERIAL_NO_OEM,
INV_CURR_USAGE.TSN_QT,
INV_CURR_USAGE.TSO_QT

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
    
INNER JOIN INV_CURR_USAGE ON
    ii.inv_no_db_id = inv_curr_usage.inv_no_db_id AND
    ii.inv_no_id = inv_curr_usage.inv_no_id

WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '32-21-00-02-1'
AND INV_CURR_USAGE.DATA_TYPE_ID = 1
'''

query_nlg_fc = f'''

SELECT 

INV_CURR_USAGE.TSN_QT,
INV_CURR_USAGE.TSO_QT

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
    
INNER JOIN INV_CURR_USAGE ON
    ii.inv_no_db_id = inv_curr_usage.inv_no_db_id AND
    ii.inv_no_id = inv_curr_usage.inv_no_id

WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '32-21-00-02-1'
AND INV_CURR_USAGE.DATA_TYPE_ID = 10
'''

query_nlg_strut = f'''
SELECT

EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
    
INNER JOIN INV_CURR_USAGE ON
    ii.inv_no_db_id = inv_curr_usage.inv_no_db_id AND
    ii.inv_no_id = inv_curr_usage.inv_no_id

WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '32-20-00-04'
AND INV_CURR_USAGE.DATA_TYPE_ID = 1


'''

query_mlg_lh_fh = f'''

SELECT 
EQP_PART_NO.PART_NO_OEM,
II.SERIAL_NO_OEM,
INV_CURR_USAGE.TSN_QT,
INV_CURR_USAGE.TSO_QT

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
    
INNER JOIN INV_CURR_USAGE ON
    ii.inv_no_db_id = inv_curr_usage.inv_no_db_id AND
    ii.inv_no_id = inv_curr_usage.inv_no_id

WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '32-11-00-02-1'
AND INV_CURR_USAGE.DATA_TYPE_ID = 1
'''

query_mlg_lh_fc = f'''

SELECT 

INV_CURR_USAGE.TSN_QT,
INV_CURR_USAGE.TSO_QT

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
    
INNER JOIN INV_CURR_USAGE ON
    ii.inv_no_db_id = inv_curr_usage.inv_no_db_id AND
    ii.inv_no_id = inv_curr_usage.inv_no_id

WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '32-11-00-02-1'
AND INV_CURR_USAGE.DATA_TYPE_ID = 10
'''

query_mlg_lh_strut = f'''
SELECT

EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
    
INNER JOIN INV_CURR_USAGE ON
    ii.inv_no_db_id = inv_curr_usage.inv_no_db_id AND
    ii.inv_no_id = inv_curr_usage.inv_no_id

WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '32-11-21-03-15-1'
AND INV_CURR_USAGE.DATA_TYPE_ID = 1

'''


query_mlg_rh_fh = f'''

SELECT 
EQP_PART_NO.PART_NO_OEM,
II.SERIAL_NO_OEM,
INV_CURR_USAGE.TSN_QT,
INV_CURR_USAGE.TSO_QT

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
    
INNER JOIN INV_CURR_USAGE ON
    ii.inv_no_db_id = inv_curr_usage.inv_no_db_id AND
    ii.inv_no_id = inv_curr_usage.inv_no_id

WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '32-11-00-02-5'
AND INV_CURR_USAGE.DATA_TYPE_ID = 1
'''

query_mlg_rh_fc = f'''

SELECT 

INV_CURR_USAGE.TSN_QT,
INV_CURR_USAGE.TSO_QT

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
    
INNER JOIN INV_CURR_USAGE ON
    ii.inv_no_db_id = inv_curr_usage.inv_no_db_id AND
    ii.inv_no_id = inv_curr_usage.inv_no_id

WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '32-11-00-02-5'
AND INV_CURR_USAGE.DATA_TYPE_ID = 10
'''

query_mlg_rh_strut = f'''
SELECT

EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
    
INNER JOIN INV_CURR_USAGE ON
    ii.inv_no_db_id = inv_curr_usage.inv_no_db_id AND
    ii.inv_no_id = inv_curr_usage.inv_no_id

WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '32-11-21-03-15-5'
AND INV_CURR_USAGE.DATA_TYPE_ID = 1

'''

query_engine_lh = f'''

SELECT 
EQP_PART_NO.PART_NO_OEM,
II.SERIAL_NO_OEM,
INV_CURR_USAGE.TSN_QT,
II.CONFIG_POS_SDESC

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
    
INNER JOIN INV_CURR_USAGE ON
    ii.inv_no_db_id = inv_curr_usage.inv_no_db_id AND
    ii.inv_no_id = inv_curr_usage.inv_no_id

WHERE ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '71-00-00-00'
AND INV_CURR_USAGE.DATA_TYPE_ID = 1
AND II.CONFIG_POS_SDESC = '71-00-00-00 (LH)'

'''

query_engine_rh = f'''

SELECT 
EQP_PART_NO.PART_NO_OEM,
II.SERIAL_NO_OEM,
INV_CURR_USAGE.TSN_QT,
II.CONFIG_POS_SDESC

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
    
INNER JOIN INV_CURR_USAGE ON
    ii.inv_no_db_id = inv_curr_usage.inv_no_db_id AND
    ii.inv_no_id = inv_curr_usage.inv_no_id

WHERE ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '71-00-00-00'
AND INV_CURR_USAGE.DATA_TYPE_ID = 1
AND II.CONFIG_POS_SDESC = '71-00-00-00 (RH)'
'''

query_engine_lh_csn = f'''


SELECT 

INV_CURR_USAGE.TSN_QT

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
    
INNER JOIN INV_CURR_USAGE ON
    ii.inv_no_db_id = inv_curr_usage.inv_no_db_id AND
    ii.inv_no_id = inv_curr_usage.inv_no_id

WHERE ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '71-00-00-00'
AND INV_CURR_USAGE.DATA_TYPE_ID = 10
AND II.CONFIG_POS_SDESC = '71-00-00-00 (LH)'

'''

query_engine_rh_csn = f'''


SELECT 

INV_CURR_USAGE.TSN_QT

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
    
INNER JOIN INV_CURR_USAGE ON
    ii.inv_no_db_id = inv_curr_usage.inv_no_db_id AND
    ii.inv_no_id = inv_curr_usage.inv_no_id

WHERE ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '71-00-00-00'
AND INV_CURR_USAGE.DATA_TYPE_ID = 10
AND II.CONFIG_POS_SDESC = '71-00-00-00 (RH)'

'''

query_apu = f'''

SELECT 
EQP_PART_NO.PART_NO_OEM,
II.SERIAL_NO_OEM,
INV_CURR_USAGE.TSN_QT

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
    
INNER JOIN INV_CURR_USAGE ON
    ii.inv_no_db_id = inv_curr_usage.inv_no_db_id AND
    ii.inv_no_id = inv_curr_usage.inv_no_id

WHERE ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '49-10-00-00'
AND INV_CURR_USAGE.DATA_TYPE_ID = 101017
'''


query_apu_csn_acyc = f'''


SELECT 

INV_CURR_USAGE.TSN_QT

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
    
INNER JOIN INV_CURR_USAGE ON
    ii.inv_no_db_id = inv_curr_usage.inv_no_db_id AND
    ii.inv_no_id = inv_curr_usage.inv_no_id

WHERE ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '49-10-00-00'
AND INV_CURR_USAGE.DATA_TYPE_ID = 101018

'''

query_fcc = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '22-11-33-03'
'''

query_fmc = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '34-61-02-02'
'''

query_autoth_c = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '22-31-10-01'
'''


query_hf = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '23-11-21-11'
'''

query_vhf = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '22-11-33-03'
'''

query_cmu = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '22-11-33-03'
'''

query_apm = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '23-27-35-05'
'''

query_cvr = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '23-71-11-01'
'''

query_fdr = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '31-31-11-05'
'''

query_dfdau = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '31-31-22-02'
'''

query_printer = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '31-33-01-01'
'''

query_dme = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '34-55-21-02'
'''

query_lrra = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '34-33-21-01'
'''

query_deu = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '31-62-11-02'
'''

query_wxr = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '34-43-41-03'
'''

query_gpws = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '34-46-01-02'
'''

query_tcas = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '34-45-01-02'
'''

query_mmr = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '34-31-42-02-6'
'''

query_xpnder = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '34-53-02-06'
'''

query_adiru = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '34-21-01-01'
'''

query_vor = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '34-51-01-02'
'''

query_du = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id 
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '31-62-11-01'
'''

query_isfd = f'''
SELECT 
EQP_MANUFACT.MANUFACT_NAME || ' ' || 'P/N ' || EQP_PART_NO.PART_NO_OEM

FROM INV_INV II 

INNER JOIN INV_INV II_AC 
    ON ii_ac.inv_no_db_id = ii.h_inv_no_db_id
    
    AND ii_ac.inv_no_id = ii.h_inv_no_id 
    AND ii_ac.inv_cond_cd NOT IN 'ARCHIVE' 
    AND ii_ac.authority_id IS NULL 
INNER JOIN INV_AC_REG 
    ON inv_ac_reg.inv_no_db_id = ii_ac.inv_no_db_id 
    AND inv_ac_reg.inv_no_id = ii_ac.inv_no_id 
INNER JOIN EQP_PART_NO ON
    ii.part_no_id = eqp_part_no.part_no_id
INNER JOIN EQP_MANUFACT ON
    EQP_PART_NO.MANUFACT_CD = EQP_MANUFACT.MANUFACT_CD
INNER JOIN EQP_BOM_PART ON
    ii.BOM_PART_DB_ID = EQP_BOM_PART.BOM_PART_DB_ID AND
    ii.BOM_PART_ID = EQP_BOM_PART.BOM_PART_ID    
WHERE ii.inv_class_cd IN 'TRK' 
AND ii.inv_cond_cd IN 'INSRV' 
AND INV_AC_REG.AC_REG_CD = '{aircraft}'
AND EQP_PART_NO.PART_NO_SDESC NOT LIKE '%SOFTWARE%'
AND EQP_BOM_PART.BOM_PART_CD = '34-24-02-01'
'''

#Connecting to Maintenix Oracle Database
dsn_tns = cx_Oracle.makedsn('maintenixdb-test.somoscopa.com', '1521', service_name='COPAT')
conn = cx_Oracle.connect(user='MX_TEST', password='MXT35T2016', dsn=dsn_tns) 

#Executing queries
c = conn.cursor()
c.execute(query_ac_info)

d = conn.cursor()
d.execute(query_tsn_fh)

e = conn.cursor()
e.execute(query_tsn_fc)

f = conn.cursor()
f.execute(query_nextcck)

g = conn.cursor()
g.execute(query_lastcck)

h = conn.cursor()
h.execute(query_nlg_fh)

i = conn.cursor()
i.execute(query_nlg_fc)

j = conn.cursor()
j.execute(query_mlg_lh_fh)

k = conn.cursor()
k.execute(query_mlg_lh_fc)

l = conn.cursor()
l.execute(query_mlg_rh_fh)

m = conn.cursor()
m.execute(query_mlg_rh_fc)

n = conn.cursor()
n.execute(query_engine_lh)

o = conn.cursor()
o.execute(query_engine_rh)

p = conn.cursor()
p.execute(query_apu)

q = conn.cursor()
q.execute(query_apu_csn_acyc)

r = conn.cursor()
r.execute(query_fcc)

s = conn.cursor()
s.execute(query_fmc)

t = conn.cursor()
t.execute(query_autoth_c)

u = conn.cursor()
u.execute(query_hf)

v = conn.cursor()
v.execute(query_vhf)

w = conn.cursor()
w.execute(query_cmu)

x = conn.cursor()
x.execute(query_apm)

y = conn.cursor()
y.execute(query_cvr)

z = conn.cursor()
z.execute(query_fdr)

aa = conn.cursor()
aa.execute(query_dfdau)

ab = conn.cursor()
ab.execute(query_printer)

ac = conn.cursor()
ac.execute(query_dme)

ad = conn.cursor()
ad.execute(query_lrra)

ae = conn.cursor()
ae.execute(query_deu)

af = conn.cursor()
af.execute(query_wxr)

ag = conn.cursor()
ag.execute(query_gpws)

ah = conn.cursor()
ah.execute(query_tcas)

ai = conn.cursor()
ai.execute(query_mmr)

aj = conn.cursor()
aj.execute(query_xpnder)

ak = conn.cursor()
ak.execute(query_adiru)

al = conn.cursor()
al.execute(query_vor)

am = conn.cursor()
am.execute(query_du)

an = conn.cursor()
an.execute(query_isfd)

#Getting aircraft ID information
for row in c: 

    ac_rg = row[0] #matricula
    man_date = row[1].strftime('%d-%b-%y') #manufactured date
    ln = row[2] #line number
    varn = row[3] #variable number
    msn = row[4] #mfg serial number
    ac_model = row[5] #tipo de 737
    if ac_model == '737-8V3':
        ac_model = 'Boeing 737-800'
    elif ac_model == '737-7V3':
        ac_model = 'Boeing 737-700'

for row in d: ac_tsn_fh = int(row[0]) #ac tsn in fh
for row in e: ac_tsn_fc = int(row[0]) #ac tsn in fc
for row in f: ac_nextcck = row[0].strftime('%d-%b-%y') #ac next c-ck due
for row in g: ac_lastcck = row[0].strftime('%d-%b-%y') #ac last c-ck due

if aircraft in ['HP-1376CMP','HP-1523CMP','HP-1532CMP','HP-1536CMP','HP-1537CMP']: #Aircraft operator
    operator = 'AEROREPUBLICA'
else:
    operator = 'COPA AIRLINES'


#NLG P/N, S/N, TSN (FH), TSO (FH)
for row in h: 
    nlg_pn = row[0]
    nlg_sn = row[1]
    nlg_tsn_fh = int(row[2])
    nlg_tso_fh = int(row[3])
for row in i: #NLG TSN (FC), TSO (FC)
    nlg_tsn_fc = int(row[0])
    nlg_tso_fc = int(row[1])

nlg_cbo = 18000 #NLG cycles between overhaul
nlg_cno = nlg_cbo - nlg_tso_fc #NLG cycles to next overhaul

if nlg_pn == 'B737-789NLG': #if NLG P/N is weird then get the strut P/N
    ha = conn.cursor()
    ha.execute(query_nlg_strut)
    for row in ha:
        nlg_pn = row[0]

#MLG LH P/N, S/N, TSN (FH), TSO (FH)  
for row in j: 
    mlg_lh_pn = row[0]
    mlg_lh_sn = row[1]
    mlg_lh_tsn_fh = int(row[2])
    mlg_lh_tso_fh = int(row[3])
for row in k: #MLG LH TSN (FC), TSO (FC)
    mlg_lh_tsn_fc = int(row[0])
    mlg_lh_tso_fc = int(row[1])

if mlg_lh_pn == 'B737-7LMG' or 'B737-89LMG': #if MLG LH P/N is weird then get the strut P/N
    ia = conn.cursor()
    ia.execute(query_mlg_lh_strut)
    for row in ia:
        mlg_lh_pn = row[0]

mlg_cbo = 21000 #MLG cycles between overhaul
mlg_lh_cno = mlg_cbo - mlg_lh_tso_fc #MLG L/H cycles to next overhaul

#MLG RH P/N, S/N, TSN (FH), TSO (FH)
for row in l: 
    mlg_rh_pn = row[0]
    mlg_rh_sn = row[1]
    mlg_rh_tsn_fh = int(row[2])
    mlg_rh_tso_fh = int(row[3])
for row in m: #MLG RH TSN (FC), TSO (FC)
    mlg_rh_tsn_fc = int(row[0])
    mlg_rh_tso_fc = int(row[1])

if mlg_rh_pn == 'B737-7RMG' or 'B737-89RMG': #if MLG RH P/N is weird then get the strut P/N
    la = conn.cursor()
    la.execute(query_mlg_rh_strut)
    for row in la:
        mlg_rh_pn = row[0]
        
mlg_cbo = 21000 #MLG cycles between overhaul
mlg_rh_cno = mlg_cbo - mlg_rh_tso_fc #MLG L/H cycles to next overhaul

#LH engine data with TSN in FH
for row in n: 
    eng_lh_model = row[0]
    eng_lh_sn = int(row[1])
    eng_lh_tsn_fh = int(row[2])

#RH engine data with TSN in FH
for row in o: 
    eng_rh_model = row[0]
    eng_rh_sn = int(row[1])
    eng_rh_tsn_fh = int(row[2])

if '22' in eng_lh_model or eng_rh_model == True: #Defining thrust rating
    thrust_rating= '22,700'
else:
    thrust_rating = '26,300'

#Getting engines CSN
na = conn.cursor()
na.execute(query_engine_lh_csn)
for row in na: eng_lh_csn = int(row[0]) #LH Engine CSN
nb = conn.cursor()
nb.execute(query_engine_rh_csn)
for row in nb: eng_rh_csn = int(row[0]) #RH Engine CSN

#Getting Engine Shop Visit Times
try:
    df_er = pd.read_excel(path_er)
    filt = (df_er['Repair Type?'] == 'Shop Visit') & (df_er['ESN'] == eng_lh_sn)
    df_er1 = df_er.loc[filt].nlargest(1,'TSN')
    eng_lh_sv_tsn = int(df_er1.values[0][5])
    eng_lh_sv_csn = int(df_er1.values[0][10])
except:
    eng_lh_sv_tsn = 0
    eng_lh_sv_csn = 0
    
try:
    filt = (df_er['Repair Type?'] == 'Shop Visit') & (df_er['ESN'] == eng_rh_sn)
    df_er1 = df_er.loc[filt].nlargest(1,'TSN')
    eng_rh_sv_tsn = int(df_er1.values[0][5])
    eng_rh_sv_csn = int(df_er1.values[0][10])
except:
    eng_rh_sv_tsn = 0
    eng_rh_sv_csn = 0

eng_lh_tslv = eng_lh_tsn_fh - eng_lh_sv_tsn
eng_lh_cslv = eng_lh_csn - eng_lh_sv_csn
eng_rh_tslv = eng_rh_tsn_fh - eng_rh_sv_tsn
eng_rh_cslv = eng_rh_csn - eng_rh_sv_csn


#APU DATA with TSN in FH (AOT) and FC (ACYC)
for row in p: 
    apu_pn = row[0] 
    apu_sn = row[1]
    apu_tsn_fh_aot = int(row[2])

for row in q: apu_csn_acyc = int(row[0]) #APU CSN

#Getting numerical digits from APU S/N for parsing
pattern_apu = re.compile(r'\d+')
matches = pattern_apu.findall(apu_sn)
apu_sn2 = int(matches[0])

#Calculating APU TSLV
try:
    df_ar = pd.read_excel(path_ar)
    filt = df_ar['S/N'] == apu_sn2
    df_ar1 = df_ar.loc[filt].nlargest(1,'Date Rem')
    apu_rmv_hours = int(df_ar1.values[0][20])
    apu_tslv = apu_tsn_fh_aot - apu_rmv_hours
except:
    apu_tslv = apu_tsn_fh_aot

# Getting avionics components information
for row in r: fcc = row[0] #FCC Info
for row in s: fmc = row[0] # FMC Info
for row in t: at_computer = row[0] # Autothrottle computer
for row in u: hf = row[0] # HF Radio
for row in v: vhf= row[0] # VHF Radio
for row in w: cmu = row[0] # CMU
for row in x: apm = row[0] # APM
for row in y: cvr = row[0] # CVR
for row in z: fdr = row[0] # FDR
for row in aa: dfdau = row[0] # DFDAU
for row in ab: printer = row[0] # Printer
for row in ac: dme = row[0] # DME
for row in ad: lrra = row[0] # LRRA
for row in ae: deu = row[0] # DEU
for row in af: wxr = row[0] # WXR Radar
for row in ag: gpws = row[0] # GPWS
for row in ah: tcas = row[0] # TCAS
for row in ai: mmr = row[0] # MMR
for row in aj: xpnder = row[0] # ATC XPNDER
for row in ak: adiru = row[0] # ADIRU
for row in al: vor = row[0] # VOR
for row in am: du = row[0] # Display Unit
for row in an: isfd = row[0] # ISFD


# Setting up output excel file

user_path = os.environ['USERPROFILE']
filename = f'Spec Sheet {aircraft} MSN {msn} ({today}).xlsx'
location = os.path.join(user_path, 'Documents',f'{filename}')

if os.path.isfile(location):#Delete if filename already exists
  os.remove(location)
else:
    pass

workbook = xlsxwriter.Workbook(location)
worksheet = workbook.add_worksheet()
worksheet.set_margins(0.25,0.25,0.4,0.4)
footer = 'Aircraft specifications are intended to be preliminary information only and must be verified prior to sale.'
worksheet.set_footer(footer)

# Setting cells formatting
cell_format1 = workbook.add_format({'bold': True, 'underline':True}) #Title format 
cell_format2 = workbook.add_format()
cell_format2.set_num_format('#,##0') #Thousand separator for numbers
cell_format2.set_align('left') #Left align
cell_format3 = workbook.add_format({'align':'right'})
cell_format4 = workbook.add_format({'bold': True,'align':'right'})
cell_format5 = workbook.add_format({'bold': True,'align':'center'})
cell_format6 = workbook.add_format()
cell_format6.set_bg_color('#D3D3D3') #Section dividers background color
cell_format6.set_border()
cell_format7 = workbook.add_format() #Text wrapping
cell_format7.set_text_wrap()
cell_format8 = workbook.add_format({'align':'left'})

cell_format1.set_left()
cell_formata4 = workbook.add_format()
cell_formata4.set_left()
##cell_formata4.set_top()
cell_borders = workbook.add_format()
cell_borders.set_left()

cell_top = workbook.add_format()
cell_top.set_top()

cell_right = workbook.add_format()
cell_right.set_right()

cell_bottom = workbook.add_format()
cell_bottom.set_bottom()

cell_g98 = workbook.add_format()
cell_g98.set_bottom()
cell_g98.set_right()



cell_format6.set_left()
worksheet.write('B4', '', cell_top)
worksheet.write('C3', '', cell_bottom)
worksheet.write('A3', '', cell_bottom)
worksheet.write('D4', '', cell_top)
worksheet.write('E4', '', cell_top)
worksheet.write('F4', '', cell_top)
worksheet.write('G3', '', cell_bottom)
worksheet.write('A53', '', cell_bottom)


worksheet.write('C4', '', cell_borders)
worksheet.write('C5', '', cell_borders)
worksheet.write('C6', '', cell_borders)
worksheet.write('C7', '', cell_borders)
worksheet.write('C8', '', cell_borders)
worksheet.write('C9', '', cell_borders)
worksheet.write('C10', '', cell_borders)
worksheet.write('C11', '', cell_borders)
worksheet.write('C12', '', cell_borders)

n=4
while n <=48:
 worksheet.write(f'G{n}', '', cell_right)
 n = n + 1

for n in ['A','B','C','D','E','F','G']:
    worksheet.write(f'{n}49','',cell_top)

for n in ['B','C','D','E','F','G']:
    worksheet.write(f'{n}54','',cell_top)

worksheet.write('G53','',cell_bottom)

worksheet.write('A99','',cell_top)



n=54
while n <=98:
 worksheet.write(f'G{n}', '', cell_right)
 n = n + 1

for n in ['A','B','C','D','E','F','G']:
    worksheet.write(f'{n}98','',cell_bottom)

n=88
while n <=98:
 worksheet.write(f'A{n}', '', cell_borders)
 n = n + 1

worksheet.write('G98','',cell_g98)

for n in ['A','B','C','D','E','F','G']:
    worksheet.write(f'{n}103','',cell_bottom)
    
n=104
while n <=148:
 worksheet.write(f'G{n}', '', cell_right)
 n = n + 1

for n in ['A','B','C','D','E','F','G']:
    worksheet.write(f'{n}149','',cell_top)

n=136
while n <=148:
 worksheet.write(f'A{n}', '', cell_borders)
 n = n + 1

cell_format_header = workbook.add_format({'bold': True}) #Header format
worksheet.set_column('A:A', 25.5)
worksheet.set_column('B:B', 15)
worksheet.set_column('C:C', 21)
worksheet.set_column('G:G', 6.9)


# Headers Page 1
worksheet.write('A1', f'OPERATOR: {operator}',cell_format_header)
worksheet.merge_range('B1:E1', 'AIRCRAFT SPEC SHEET', cell_format5)
worksheet.merge_range('B2:E2', f'MSN: {msn}', cell_format5)
worksheet.merge_range('F1:G1', f'DATE: {today}', cell_format4)
worksheet.write('G2', '1 of 3',cell_format4)

# Aircraft Data
worksheet.write('A4', 'AIRFRAME',cell_format1)
worksheet.write('A5', 'Model......................................................................................',cell_formata4)
worksheet.write('A6','Registry......................................................................................',cell_borders)
worksheet.write('A7','Variable Number......................................................................................',cell_borders)
worksheet.write('A8','Serial Number......................................................................................',cell_borders)
worksheet.write('A9','Line Number......................................................................................',cell_borders)
worksheet.write('A10', 'Manufacturing Date......................................................................................',cell_borders)
worksheet.write('A11', f'Flight Hours ({today})......................................................................................',cell_borders)
worksheet.write('A12', f'Flight Cycles ({today})......................................................................................',cell_borders)
worksheet.write('B5', ac_model)
worksheet.write('B6', ac_rg)
worksheet.write('B7', varn)
worksheet.write('B8', msn)
worksheet.write('B9', ln)
worksheet.write('B10', man_date)
worksheet.write('B11', ac_tsn_fh, cell_format2)
worksheet.write('B12', ac_tsn_fc, cell_format2)
worksheet.write('A13', 'Maximum Taxi Weight......................................................................................',cell_borders)
worksheet.write('A14', 'Maximum Takeoff Weight......................................................................................',cell_borders)
worksheet.write('A15', 'Maximum Landing Weight......................................................................................',cell_borders)
worksheet.write('A16', 'Maximum Zero-Fuel Weight......................................................................................',cell_borders)
worksheet.write('A17', 'Noise Category......................................................................................',cell_borders)
worksheet.write('A18', 'Landing Category Approval......................................................................................',cell_borders)
worksheet.write('A19', 'Short Field Performance......................................................................................',cell_borders)
worksheet.write('B13', mtw)
worksheet.write('B14', mtow)
worksheet.write('B15', mlw)
worksheet.write('B16', mzfw)
worksheet.write('B17', noise_cat)
worksheet.write('B18', cat_status)
worksheet.write('B19', sfp)

# Maintenance Program
worksheet.write('C13', 'MAINTENANCE PROGRAM', cell_format1)
worksheet.write('C14', 'Last C-Check On:', cell_borders)
worksheet.write('C15', 'Next C-Check Due:', cell_borders)
worksheet.write('C16', 'C-Checks every 3 years. First C-CK at 5 years', cell_borders)
worksheet.write('C17', 'A-Checks every 120 days', cell_borders)
worksheet.write('C18', 'Engines under MCPH and Trend Monitoring', cell_borders)
worksheet.write('C19', 'APU under PBH', cell_borders)
try:
    worksheet.write('D14', ac_lastcck)
except NameError:
    worksheet.write('D14', (datetime.strptime(ac_nextcck,'%d-%b-%y') - timedelta(days=1095)).strftime('%d-%b-%y'))


worksheet.write('D15', ac_nextcck)

#Setting aircraft photo
##image_path = f'C:/Users/Giancarlo/AppData/Local/Programs/Python/Python38-32/Aircraft Photos/{aircraft}.jpg'
image_path = r'//ATO-DFS-03/WorkfilesTUM/MANTO/DISCOM/ENGINEER/REFERENCE/SpecSheetMaker737NG/Aircraft Photos/'+f'{aircraft}.jpg'
worksheet.insert_image('C4', image_path,{'object_position': 3, 'y_offset': 1})

worksheet.merge_range('A20:G20', '', cell_format6) #Section divisor

#L/H Engine Data
worksheet.write('A21', 'ENGINE L/H', cell_format1)
worksheet.write('A22', 'Type......................................................................................',cell_borders)
worksheet.write('A23', 'Thrust Rating (lb)......................................................................................',cell_borders)
worksheet.write('A24', 'Serial Number......................................................................................',cell_borders)
worksheet.write('A25', f'TSN ({today})......................................................................................',cell_borders)
worksheet.write('A26', f'CSN ({today})......................................................................................',cell_borders)
worksheet.write('A27', f'TSLSV ({today})......................................................................................',cell_borders)
worksheet.write('A28', f'CSLSV ({today})......................................................................................',cell_borders)
worksheet.write('B22', eng_lh_model)
worksheet.write('B23', thrust_rating)
worksheet.write('B24', eng_lh_sn,cell_format8)
worksheet.write('B25', eng_lh_tsn_fh,cell_format2)
worksheet.write('B26', eng_lh_csn, cell_format2)
worksheet.write('B27', eng_lh_tslv, cell_format2)
worksheet.write('B28', eng_lh_cslv, cell_format2)

#R/H Engine Data
worksheet.write('C21', 'ENGINE R/H', cell_format1)
worksheet.write('C22', 'Type......................................................................................', cell_borders)
worksheet.write('C23', 'Thrust Rating (lb)......................................................................................', cell_borders)
worksheet.write('C24', 'Serial Number......................................................................................', cell_borders)
worksheet.write('C25', f'TSN ({today})......................................................................................', cell_borders)
worksheet.write('C26', f'CSN ({today})......................................................................................', cell_borders)
worksheet.write('C27', f'TSLSV ({today})......................................................................................', cell_borders)
worksheet.write('C28', f'CSLSV ({today})......................................................................................', cell_borders)
worksheet.write('D22', eng_rh_model)
worksheet.write('D23', thrust_rating)
worksheet.write('D24', eng_rh_sn,cell_format8)
worksheet.write('D25', eng_rh_tsn_fh,cell_format2)
worksheet.write('D26', eng_rh_csn, cell_format2)
worksheet.write('D27', eng_rh_tslv, cell_format2)
worksheet.write('D28', eng_rh_cslv, cell_format2)

worksheet.merge_range('A29:G29', '', cell_format6) #Section divisor

#APU Data
worksheet.write('A30', 'APU', cell_format1)
worksheet.write('A31', 'Type......................................................................................',cell_borders)
worksheet.write('A32', 'Serial Number......................................................................................',cell_borders)
worksheet.write('A33', f'TSN ({today})......................................................................................',cell_borders)
worksheet.write('A34', f'CSN ({today})......................................................................................',cell_borders)
worksheet.write('A35', f'TSLSV ({today})......................................................................................',cell_borders)
worksheet.write('A36', '',cell_borders)
worksheet.write('A37', '',cell_borders)
worksheet.write('A38', '',cell_borders)

worksheet.write('B31', 'Honeywell 131-9B')
worksheet.write('B32', apu_sn)
worksheet.write('B33', apu_tsn_fh_aot, cell_format2)
worksheet.write('B34', apu_csn_acyc, cell_format2)
worksheet.write('B35', apu_tslv, cell_format2)


#NLG Data

worksheet.write('C30', 'NOSE LANDING GEAR', cell_format1)
worksheet.write('C31', 'Part Number (Goodrich)......................................................................................', cell_borders)
worksheet.write('C32', 'Serial Number......................................................................................', cell_borders)
worksheet.write('C33', f'TSN ({today})......................................................................................', cell_borders)
worksheet.write('C34', f'CSN ({today})......................................................................................', cell_borders)
worksheet.write('C35', f'TSO ({today})......................................................................................', cell_borders)
worksheet.write('C36', f'CSO ({today})......................................................................................', cell_borders)
worksheet.write('C37', 'CBO......................................................................................', cell_borders)
worksheet.write('C38', 'Cycles to Next Overhaul......................................................................................', cell_borders)

worksheet.write('D31', nlg_pn)
worksheet.write('D32', nlg_sn)
worksheet.write('D33', nlg_tsn_fh, cell_format2)
worksheet.write('D34', nlg_tsn_fc, cell_format2)
worksheet.write('D35', nlg_tso_fh, cell_format2)
worksheet.write('D36', nlg_tso_fc, cell_format2)
worksheet.write('D37', nlg_cbo, cell_format2)
worksheet.write('D38', nlg_cno, cell_format2)

worksheet.merge_range('A39:G39', '', cell_format6) # Section divisor

# MLG L/H Data
worksheet.write('A40', 'MAIN LANDING GEAR L/H', cell_format1)
worksheet.write('A41', 'Part Number (Goodrich)......................................................................................',cell_borders)
worksheet.write('A42', 'Serial Number......................................................................................',cell_borders)
worksheet.write('A43', f'TSN ({today})......................................................................................',cell_borders)
worksheet.write('A44', f'CSN ({today})......................................................................................',cell_borders)
worksheet.write('A45', f'TSO ({today})......................................................................................',cell_borders)
worksheet.write('A46', f'CSO ({today})......................................................................................',cell_borders)
worksheet.write('A47', 'CBO......................................................................................',cell_borders)
worksheet.write('A48', 'Cycles to Next Overhaul......................................................................................',cell_borders)

worksheet.write('B41', mlg_lh_pn)
worksheet.write('B42', mlg_lh_sn)
worksheet.write('B43', mlg_lh_tsn_fh, cell_format2)
worksheet.write('B44', mlg_lh_tsn_fc, cell_format2)
worksheet.write('B45', mlg_lh_tso_fh, cell_format2)
worksheet.write('B46', mlg_lh_tso_fc, cell_format2)
worksheet.write('B47', mlg_cbo, cell_format2)
worksheet.write('B48', mlg_lh_cno, cell_format2)

# MLG R/H Data
worksheet.write('C40', 'MAIN LANDING GEAR R/H', cell_format1)
worksheet.write('C41', 'Part Number (Goodrich)......................................................................................', cell_borders)
worksheet.write('C42', 'Serial Number......................................................................................', cell_borders)
worksheet.write('C43', f'TSN ({today})......................................................................................', cell_borders)
worksheet.write('C44', f'CSN ({today})......................................................................................', cell_borders)
worksheet.write('C45', f'TSO ({today})......................................................................................', cell_borders)
worksheet.write('C46', f'CSO ({today})......................................................................................', cell_borders)
worksheet.write('C47', 'CBO......................................................................................', cell_borders)
worksheet.write('C48', 'Cycles to Next Overhaul......................................................................................', cell_borders)

worksheet.write('D41', mlg_rh_pn)
worksheet.write('D42', mlg_rh_sn)
worksheet.write('D43', mlg_rh_tsn_fh, cell_format2)
worksheet.write('D44', mlg_rh_tsn_fc, cell_format2)
worksheet.write('D45', mlg_rh_tso_fh, cell_format2)
worksheet.write('D46', mlg_rh_tso_fc, cell_format2)
worksheet.write('D47', mlg_cbo, cell_format2)
worksheet.write('D48', mlg_rh_cno, cell_format2)

# Headers Page 2
worksheet.write('A51', f'OPERATOR: {operator}',cell_format_header)
worksheet.merge_range('B51:E51', 'AIRCRAFT SPEC SHEET', cell_format5)
worksheet.merge_range('B52:E52', f'MSN: {msn}', cell_format5)
worksheet.merge_range('F51:G51', f'DATE: {today}', cell_format4)
worksheet.write('G52', '2 of 3',cell_format4)

# Avionics Data
worksheet.write('A54', 'AVIONICS',cell_format1)
worksheet.write('A55', 'Flight Control Computer...............................................',cell_borders)
worksheet.write('A56', 'Flight Management Computer...............................................',cell_borders)
worksheet.write('A57', 'Autothrottle Computer...............................................',cell_borders)
worksheet.write('A58', 'Iridium Satcom......................................................................',cell_borders)
worksheet.write('A59', 'HF Radios Quantity...............................................................',cell_borders)
worksheet.write('A60', 'HF Comm Transceiver...............................................',cell_borders)
worksheet.write('A61', 'HFDL Enabled......................................................................',cell_borders)
worksheet.write('A62', 'VHF Radios(Qty=3)......................................................................',cell_borders)
worksheet.write('A63', 'CMU......................................................................................',cell_borders)
worksheet.write('A64', 'HFDL Override Switch......................................................................',cell_borders)
worksheet.write('A65', 'APM Aircraft Personality Module...............................................',cell_borders)
worksheet.write('A66', 'Battery......................................................................................',cell_borders)
worksheet.write('A67', 'CVR (Cockpit Voice Recorder)...............................................',cell_borders)
worksheet.write('A68', 'FDR (Flight Data Recorder)...............................................',cell_borders)
worksheet.write('A69', 'DFDAU (Digital Flight Data Acquisition Unit)...............................................',cell_borders)
worksheet.write('A70', 'DFDAU Mandatory Software...............................................',cell_borders)
worksheet.write('A71', 'RIPS (Recorder Independent Power Supply)...............................................',cell_borders)
worksheet.write('A72', 'Printer (P8 Panel)......................................................................',cell_borders)
worksheet.write('A73', 'DME (Distance Measuring Equipment)...............................................',cell_borders)
worksheet.write('A74', 'LRRA (Low Range Radio Altimeter)...............................................',cell_borders)
worksheet.write('A75', 'DEU (Display Electronics Unit)...............................................',cell_borders)
worksheet.write('A76', 'Weather Radar Receiver/Transceiver...............................................',cell_borders)
worksheet.write('A77', 'Ground Proximity Warning Computer...............................................',cell_borders)
worksheet.write('A78', 'TCAS Computer......................................................................',cell_borders)
worksheet.write('A79', 'TCAS Software......................................................................',cell_borders)
worksheet.write('A80', 'Multi-Mode Receiver......................................................................',cell_borders)
worksheet.write('A81', 'ATC Transponder......................................................................',cell_borders)
worksheet.write('A82', 'ADIRU (Air Data Inertial Reference Unit)...............................................',cell_borders)
worksheet.write('A83', 'VOR/Marker Beacon Receiver...............................................',cell_borders)
worksheet.write('A84', 'ADF Receiver......................................................................',cell_borders)
worksheet.write('A85', 'Display Unit......................................................................',cell_borders)
worksheet.write('A86', 'Integrated Standby Flight Display...............................................',cell_borders)
worksheet.write('A87', 'In-Flight Entertainment (IFE)...............................................',cell_borders)

worksheet.write('C55', fcc)
worksheet.write('C56', fmc)
try: 
    worksheet.write('C57', at_computer)
except NameError:
    worksheet.write('C57', 'Incorporated in FCC')
worksheet.write('C58', 'YES')
worksheet.write('C59', '1')
worksheet.write('C60', hf)
worksheet.write('C61', 'YES')
worksheet.write('C62', vhf)
worksheet.write('C63', cmu)
worksheet.write('C64', 'HFDL Override Switch') #verify
worksheet.write('C65', apm)
worksheet.write('C66', batt)
worksheet.write('C67', cvr)
worksheet.write('C68', fdr)
worksheet.write('C69', dfdau)
worksheet.write('C70', dfdau_sw)
worksheet.write('C71', rips)
worksheet.write('C72', printer)
worksheet.write('C73', dme)
worksheet.write('C74', lrra)
worksheet.write('C75', deu)
worksheet.write('C76', wxr)
worksheet.write('C77', gpws)
worksheet.write('C78', tcas)
worksheet.write('C79', tcas_sw)
worksheet.write('C80', mmr)
worksheet.write('C81', xpnder)
worksheet.write('C82', adiru)
worksheet.write('C83', vor)
worksheet.write('C84', 'NO')
worksheet.write('C85', du)
worksheet.write('C86', isfd)
worksheet.write('C87', ife)

# Headers Page 3
worksheet.write('A101', f'OPERATOR: {operator}',cell_format_header)
worksheet.merge_range('B101:E101', 'AIRCRAFT SPEC SHEET', cell_format5)
worksheet.merge_range('B102:E102', f'MSN: {msn}', cell_format5)
worksheet.merge_range('F101:G101', f'DATE: {today}', cell_format4)
worksheet.write('G102', '3 of 3',cell_format4)

# Interiors Data

worksheet.write('A104', 'INTERIORS',cell_format1)
worksheet.write('A105', 'BSI Interior.................................................................',cell_borders)
worksheet.write('A106', 'Passengers BC / TC...................................',cell_borders)
worksheet.write('A107', 'Seats Manuf. BC / TC.................................................',cell_borders)
worksheet.write('A108', 'Seats pitch BC / TC................................................',cell_borders)
worksheet.write('A109', 'Seats recline BC / TC................................................',cell_borders)
worksheet.write('A110', 'Galley G1..................................................',cell_borders)
worksheet.write('A111', 'Galley G2...................................................',cell_borders)
worksheet.write('A112', 'Galley G4B...................................................',cell_borders)
worksheet.write('A113', 'Galley G7.....................................................',cell_borders)
worksheet.write('A114', 'Lavatory Configuration.................................................',cell_borders)
worksheet.write('A115', 'Lavatory Manufacturer.................................................',cell_borders)
worksheet.write('A116', 'Ovens G2.........................................................',cell_borders)
worksheet.write('A117', 'Ovens G4...........................................................',cell_borders)
worksheet.write('A118', 'Observer Seats...............................................',cell_borders)
worksheet.write('A119', 'Escape Slides.........................................................',cell_borders)
worksheet.write('A120', 'ELT Fixed / Portable................................................',cell_borders)

worksheet.write('B105', bsi)
worksheet.write('B106', pax_config)
worksheet.write('B107', seats_mfg)
worksheet.write('B108', seats_pitch)
worksheet.write('B109', seats_recline)
worksheet.write('B110', g1)
worksheet.write('B111', g2)
worksheet.write('B112', g4b)
worksheet.write('B113', g7)
worksheet.write('B114', lav_config)
worksheet.write('B115', lav_mfg)
worksheet.write('B116', oven_g2)
worksheet.write('B117', oven_g4)
worksheet.write('B118', '1')
worksheet.write('B119', slides)
worksheet.write('B120', f'{elt_fixed}/{elt_port}')

worksheet.merge_range('A121:G121', '', cell_format6) # Section divisor


# Systems Data
worksheet.write('A122', 'SYSTEMS',cell_format1)
worksheet.write('A123', 'Auxiliary Fuel Tanks.................................................................',cell_borders)
worksheet.write('A124', 'Main Brakes Type...................................',cell_borders)
worksheet.write('A125', 'Main Brakes Manuf. & P/N.................................................',cell_borders)
worksheet.write('A126', 'Main Wheels Manuf. & PN................................................',cell_borders)
worksheet.write('A127', 'Switch-Dispatch w/ LG Down................................................',cell_borders)
worksheet.write('A128', '22-min Chemical O2 Generators..................................................',cell_borders)
worksheet.write('A129', 'First Obs. Full-Face Mask..................................................',cell_borders)
worksheet.write('A130', 'Potable Water Capacity..................................................',cell_borders) 
worksheet.write('A131', 'Nitrogen Generation System..................................................',cell_borders)


worksheet.write('B123', aft)
worksheet.write('B124', brakes_type)
worksheet.write('B125', brakes_mfgpn)
worksheet.write('B126', wheels_mfgpn)
worksheet.write('B127', lg_switch)
worksheet.write('B128', oxy_gen)
worksheet.write('B129', obs_mask)
worksheet.write('B130', water_tank)
worksheet.write('B131', ngs)

worksheet.merge_range('A132:G132', '', cell_format6) #Section divisor

# Structures Data
worksheet.write('A133', 'STRUCTURES',cell_format1)
worksheet.write('A134', 'Enhanced Security Cockpit Door.................................................................',cell_borders)
worksheet.write('A135', 'Winglets...................................',cell_borders)

worksheet.write('B134', fd_door)
worksheet.write('B135', winglets)

# Closing workbook
workbook.close()

# Opening file
os.system(f'"{location}"')



# todo cHECK 1821CMP APU
# todo Review c-ck dates, for example look at 1827
# todo Report Missing variable number for HP-1854CMP
# todo GIT version control
# todo Add avionics models?
# todo Arreglar potable water capacity in config matrix
# todo Todos los trenes son goodrich?
# todo xls writer vba to update values in config matrix file



