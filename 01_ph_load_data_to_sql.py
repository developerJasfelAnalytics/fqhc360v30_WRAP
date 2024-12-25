
import pandas as pd

import string
import csv
import configparser
import plotly.express as px
import plotly.io as pio
pio.templates.default = "simple_white"
pio.renderers.default = "browser"


config = configparser.ConfigParser()        # get config vars from file
config.read('C:\\Users\\gmann\\Documents\\Dev\\fqhc360v30_WRAP\\Project_HOPE\\config.ini')

# --------------------------------------------------------------
# get configuration variables from config file
health_center = config["site"]["health_center"]
data_src_dir = config['DIR']['data_src_dir']                        # set path to dir containing files
workdir = config['DIR']['work_dir']
fqhc_billing_number = config['site']['fqhc_billing_number']
remote_engine = config['database']['con']

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

# use the remote engine to write the data to the database
dfd.to_sql('ph_charges_detail', remote_engine, if_exists='replace', index=False)
