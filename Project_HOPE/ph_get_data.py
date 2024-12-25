
# the report used to pull data appears to use the 'Charge Detail' report in
# Greenway. The report is filtered by the following:
#   'Chg Cv1 Plan Code' in ['HZNJ', 'AMGP', 'UHCCP', 'AETBH', 'WELLCAID']
#   'Amount Charge' > 0
#   'Amount Payment' != Empty
#   'Record Type Desc' in ['Charge', 'Payment Assignment']
#   'Date Svc From' >= 2024-01-01 and 'Date Svc From' <= 2024-12-31
#
# Transform the data to match the Wraparound Report
# this script will be used to create the data input format needed for the wraparound
# report.
import string
import csv
import configparser
import plotly.express as px
import plotly.io as pio
pio.templates.default = "simple_white"
pio.renderers.default = "browser"
from wrap_common_functions.PH_process_definitions import *
from wrap_common_functions.wrap_spreadsheet_build_functions import *

config = configparser.ConfigParser()        # get config vars from file
config.read('C:\\Users\\gmann\\Documents\\Dev\\fqhc360v30_WRAP\\Project_HOPE\\config.ini')

# --------------------------------------------------------------
# get configuration variables from config file
health_center = config["site"]["health_center"]
data_src_dir = config['DIR']['data_src_dir']                        # set path to dir containing files
workdir = config['DIR']['work_dir']
fqhc_billing_number = config['site']['fqhc_billing_number']
# --------------------------------------------------------------

# read and concatenate the data from the Greenway report. there are three files with
# the same format. the data will be concatenated into one dataframe.
file_list = [
    'PH_Charges_Detail_2024.xlsx',
    'PH_Charges_Detail_2023.xlsx',
    'PH_Charges_Detail_2022.xlsx'
]

dfd = pd.DataFrame()
for i in range(len(file_list)):
    print('Reading:', file_list[i])
    dfd = pd.concat([dfd, pd.read_excel(data_src_dir + file_list[i])], ignore_index=True)


# dfd = pd.read_excel(src_dir + 'PH_Charges_Detail_2024.xlsx')        # get the data. this will eventually be replaced with a call to postgresql
dfd['Pat Name'] = dfd['Pat First Name'] + ' ' + dfd['Pat Last Name']
dfd['Credited Prov Name'] = dfd['Credited Prov Name'].str.strip()   # remove trailing spaces from 'Credited Prov Name'
dfd['Credited Prov Pstn Name'] = dfd['Credited Prov Pstn Name'].str.strip()
dfd['comment'] = ''

# get only columns needed for the wraparound report
dfd = dfd[[
    'Amount Charge',
    'Chg Cv1 Claim Member ID',
    'Chg Cv1 Plan Code',
    'Chg Cv1 Grp Number',
    'Chg Claim 1 Claim Number',
    'Credited Prov Pstn Name',
    'Credited Prov Name',
    'Date Post Pmt',
    'Date Svc From',
    'Encounter Nbr',
    'Pat DOB',
    'Pat Name',
    'Procedure Code',
    'Record Type Desc'
]]

# let's map out what needs to be done here.
# 1. after we get the data from greenway, we need to do the following:
#    a. filter out encounters that do not meet the wraparound report criteria.
#       a1. 'Chg Cv1 Plan Code' in ['HZNJ', 'AMGP', 'UHCCP', 'AETBH', 'WELLCAID']
#    b. transform the data to match the wraparound report
#    c. create the Execl spreadsheet and send to finance team.

year = '2023'
quarter = 'Q4'
report_type = 'RECON'
# report_type = 'RECON'
pay_rate = 219.83
horizon_capitation = 0
show_encounters_with_no_claim_date = False

# using the quarter and year, get the start_date, end_date, sheets, quarter_with_year, and months
start_date, end_date, sheets, quarter_with_year, months = get_quarter_info(quarter, year)

print("Start Date:", start_date)
print("End Date:  ", end_date)
print("Sheets:    ", sheets)
print("Quarter:   ", quarter_with_year)
print("Months:    ", months)

data_detail = dfd.copy()
data_detail['date_of_service'] = pd.to_datetime(data_detail['Date Svc From'])       # convert 'Date Svc From' to datetime
data_detail['service_month'] = data_detail['date_of_service'].dt.strftime('%Y-%m')  # get the month of service

# filter out encounters
data_detail = data_detail[data_detail['service_month'].isin(months)].copy()                 # get rows where 'service_month' is in months
data_detail = data_detail[data_detail['Chg Cv1 Plan Code'].isin(ph_hmo_dict.keys())].copy()    # get rows where 'Chg Cv1 Plan Code' is in hmo_dict
data_detail = data_detail[data_detail['Credited Prov Pstn Name'].isin(ph_valid_service_types)].copy()  # get rows where 'Chg Cv1 Plan Code' is in valid_service_types
data_detail.rename(columns={'Credited Prov Pstn Name': 'service_type'}, inplace=True)      # this maps service type to CPT4 codes

data_detail.rename(columns={'Chg Cv1 Plan Code': 'hmo_name'}, inplace=True)              # rename 'Chg Cv1 Plan Code' to 'hmo_name'
data_detail.replace({'hmo_name': ph_hmo_dict}, inplace=True)                             # replace 'hmo_name' with values from hmo_dict
data_detail['hmo_name'] = data_detail['hmo_name'].astype(str)                           # convert hmo_name to 3 digit string
data_detail['hmo_name'] = data_detail['hmo_name'].str.zfill(3)                          # pad hmo_name with leading zeros

data_detail['fqhc_billing_number'] = fqhc_billing_number
data_detail.rename(columns={'Date Post Pmt': 'claim_payment_date'}, inplace=True)       # rename 'Date Post Pmt' to 'claim_payment_date'
data_detail.rename(columns={'Encounter Nbr': 'enc_nbr'}, inplace=True)                  # rename 'Encounter Nbr' to 'enc_nbr'
data_detail.rename(columns={'Amount Charge': 'claim_payment_amount'}, inplace=True)     # rename 'Amount Charge' to 'claim_payment_amount'

missing_dates = data_detail[data_detail['claim_payment_date'] == ''].copy()             # get rows where claim_payment_date is empty
# data_detail = data_detail[data_detail['claim_payment_date'].notnull()].copy()               # get rows where claim_payment_date is not empty

data_detail = data_detail.reset_index(drop=True)            # reindex data_detail
data_detail['encounter'] = 1                                # create encounter column and set to 1

# sort data_detail by 'CLAIM_ID' and 'date_of_service'
data_detail = data_detail.sort_values(by=['enc_nbr', 'date_of_service'], ascending=True).reset_index(drop=True)

# print the number of rows in data_detail
print('Number of claim lines in data_detail:', '{:,}'.format(data_detail.shape[0]))

# ----------------------------------------------------------------------------------------
# at this point, we have the data we need to create the wraparound report.
# everything from here on out is to transform the data to match the wraparound report.
# we will use common functions from the wrap_common_functions module to do this.
# ----------------------------------------------------------------------------------------
data_detail = update_encounter_column(data_detail)                              # update the 'encounter' column

data1, data2, data3 = process_data_by_month(data_detail, months)               # process data by month

# if service type is not in the 'Service Type' column add it
# column_values should be ph_valid_service_types without the 'Unknown' value
column_values = ph_valid_service_types.copy()
if 'Unknown' in column_values:                      # remove 'Unknown' from column_values
    column_values.remove('Unknown')
for column_value in column_values:
    data1, data2, data3 = add_missing_service_types(column_value, data1, data2, data3)

data1, data2, data3 = process_core_services(data1, data2, data3)            # process core services

###############################################################################################
# Spreadsheet build section. this should eventually be moved to a separate function
# build the spreadsheet
###############################################################################################
inum_encounters = 0
itotal_payment = inum_encounters * pay_rate
imanaged_care_receipts = 0
ivaccine_receipts = 0
idifference = itotal_payment - imanaged_care_receipts - ivaccine_receipts


rnum_encounters = data_detail['enc_nbr'].nunique()              # get unique encounters
print('# of unique encounters:', '{:,}'.format(rnum_encounters))

# create dataframe for Page 1
rtotal_payment = rnum_encounters * pay_rate

rmanaged_care_receipts = data_detail['claim_payment_amount']                # get claim_payment_amount
rmanaged_care_receipts = rmanaged_care_receipts.astype(float)               # convert to float
rmanaged_care_receipts = rmanaged_care_receipts.sum()
rmanaged_care_receipts = abs(rmanaged_care_receipts)

rvaccine_receipts = 0
rdifference = rtotal_payment - rmanaged_care_receipts - rvaccine_receipts
ramount_due = abs(rvaccine_receipts - rmanaged_care_receipts)

# create dataframe for Page 1
page1_data = [
    ['A', 'Total Encounters', inum_encounters, rnum_encounters ],
    ['B', 'Medicaid PPS Rate', pay_rate, pay_rate],
    ['C', 'Total Payment', itotal_payment, rtotal_payment],
    ['D', 'Managed Care Receipts', imanaged_care_receipts, rmanaged_care_receipts],
    ['E', 'Vaccine Receipts', ivaccine_receipts, rvaccine_receipts],
    ['F', 'Difference', idifference, rdifference],
    ['G', 'Amount Due', 0, ramount_due]
]
page1_data = pd.DataFrame(page1_data, columns=['', 'Metric', 'Initial', 'Reconciliation'])

# -----------------------------------
# put together schedule B
# -----------------------------------
schedB = data_detail.copy()

# month names is used in the schedules
month_names = ['', '', '']                                  # create empty list for month names
month_names[0] = pd.to_datetime(months[0]).strftime('%b')
month_names[1] = pd.to_datetime(months[1]).strftime('%b')
month_names[2] = pd.to_datetime(months[2]).strftime('%b')

schedB1_sum, schedB2_sum, schedB3_sum = process_schedule_B(schedB, months, horizon_capitation)


# # format claim_payment_date to mm/dd/yyyy
# data_detail['claim_payment_date'] = data_detail['claim_payment_date'].dt.strftime('%m/%d/%Y')
data_detail['date_of_service'] = data_detail['date_of_service'].dt.strftime('%m/%d/%Y')
data_detail['enc_nbr'] = data_detail['enc_nbr'].astype(str).str[-6:]            # get last 6 digits of enc_nbr
data_detail['comment'] = ''

# rename columns to reflect WRAP report column names recognized by the state
data_detail = data_detail.rename(columns={'fqhc_billing_number': 'BILLING_PROV_ NO',
                                          'Chg Cv1 Claim Member ID': 'MEDICAID_RCP_ID_NO',
                                          'Pat Name': 'MEDICAID_MEDICAID_RCP_FULL_NAME',
                                          'Pat DOB': 'MEDICAID_RCP_BIRTH_DATE',
                                          'hmo_name': 'HMO_NAME',
                                          'Chg Cv1 Grp Number': 'MEDICAID_RCP_HMO_ASSIGNED_ID',
                                          'date_of_service': 'CLM_SVC_DTE',
                                          'Procedure Code': 'CLM_CPT_CDE',
                                          'service_type': 'SERVICE_TYPE',
                                          'claim_payment_date': 'CLM_PMT_AMT_DATE',
                                          'encounter': 'ENCOUNTER',
                                          'claim_payment_amount': 'CLM_PMT_AMT',
                                          'enc_nbr': 'Claim ID',
                                          'comment': 'COMMENT'
                                          })

# convert 'MEDICAID_RCP_BIRTH_DATE' to mm/dd/yyyy
data_detail['MEDICAID_RCP_BIRTH_DATE'] = pd.to_datetime(data_detail['MEDICAID_RCP_BIRTH_DATE'])
data_detail['MEDICAID_RCP_BIRTH_DATE'] = data_detail['MEDICAID_RCP_BIRTH_DATE'].dt.strftime('%m/%d/%Y')
data_detail = data_detail[[
    'BILLING_PROV_ NO', 'MEDICAID_RCP_ID_NO', 'MEDICAID_MEDICAID_RCP_FULL_NAME', 'MEDICAID_RCP_BIRTH_DATE',
    'HMO_NAME', 'MEDICAID_RCP_HMO_ASSIGNED_ID', 'CLM_SVC_DTE', 'CLM_CPT_CDE', 'SERVICE_TYPE', 'CLM_PMT_AMT_DATE',
    'ENCOUNTER', 'CLM_PMT_AMT', 'Claim ID', 'COMMENT'
]]

# make the following columns strings: MEDICAID_MEDICAID_RCP_FULL_NAME, HMO_NAME,
# MEDICAID_RCP_ID_NO
data_detail['MEDICAID_MEDICAID_RCP_FULL_NAME'] = data_detail['MEDICAID_MEDICAID_RCP_FULL_NAME'].astype(str)
data_detail['HMO_NAME'] = data_detail['HMO_NAME'].astype(str)
data_detail['MEDICAID_RCP_ID_NO'] = data_detail['MEDICAID_RCP_ID_NO'].astype(str)
data_detail['MEDICAID_RCP_HMO_ASSIGNED_ID'] = data_detail['MEDICAID_RCP_HMO_ASSIGNED_ID'].astype(str)

# create data sheet from data_detail
data_detail_sheet_name = 'Data Detail'

# ----------------------------------------------------------------------
# Create dataframe for schedule C and use schedule C for D, E, F as well
# ----------------------------------------------------------------------
schedC_data = [['', '', '', '', '', '', '', 0] for _ in range(26)]          # create empty DataFrame
column_names = list(string.ascii_uppercase[:len(schedC_data[0])])       # get column names
schedC_data = pd.DataFrame(schedC_data, columns=column_names)             # create DataFrame

# ----------------------------------------------------
# build spreadsheet with page1, the detail data and the schedules
# ----------------------------------------------------
build_spreadsheet_and_schedules(health_center, fqhc_billing_number, ph_hmo_dict, year, workdir, quarter_with_year, report_type,
                                page1_data, data_detail,
                                data1, data2, data3,
                                schedB1_sum, schedB2_sum, schedB3_sum, schedC_data,
                                sheets, month_names)

