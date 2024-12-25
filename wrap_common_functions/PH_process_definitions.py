import pandas as pd

# create a dictionary to map HMO with their respective codes
ph_hmo_dict = {
    'AMGP':     '001',
    'UHCCP':    '002',
    'HZNJ':     '003',
    'AETBH':    '004',
    'WELLCAID': '005'
}

ph_provider_services = {
    'Al-Hilli, Rula MD': 'Physician',
    'Buch, Deepak MD': 'Physician',
    'Farestad, Jennifer L AGNP': 'Nurse Practitioner',
    "O'Connor, Veronica APN": 'Nurse Practitioner',
    'Ogunleye, Adetutu Mojisola S AGNP': 'Nurse Practitioner',
    'Satterfield, Ashley Dani APN': 'Nurse Practitioner',
    'Stringer, Daniel APN': 'Nurse Practitioner',
    'Payton, Niema RN': 'Nursing',
    'Speight, Tracey': 'Nursing',
}
ph_provider_services2 = {
    'Farestad, Jennifer L AGNP': 'Nurse Practitioner',
    'Satterfield, Ashley Dani APN': 'Nurse Practitioner',
    'Al-Hilli, Rula MD': 'Doctor',
    'O\'Connor, Veronica APN': 'Nurse Practitioner',
    'Ogunleye, Adetutu Mojisola S AGNP': 'Nurse Practitioner',
    'Payton, Niema RN': 'Registered Nurse',
    'Stringer, Daniel APN': 'Nurse Practitioner',
    'Speight, Tracey': 'Registered Nurse',
    'Ruiz, Alexza': 'Other',
    'Buch, Deepak MD': 'Doctor',
    'Sherrod, Jacqui LPN': 'Practical Nurse',
    'Powell, Molly DNP': 'Nurse Practitioner',
    'Thompson, Adrian R LCSW': 'Licensed Clinical Social Worker',
    'Smith, Zachariah APN': 'Nurse Practitioner',
    'Milby, Philip LCSW': 'Licensed Clinical Social Worker',
    'Thompson, Andrea N LCSW': 'Licensed Clinical Social Worker',
    'Gilchrist, Paris LSW': 'Licensed Social Worker',
    'Jackson, Kathleen APN': 'Nurse Practitioner',
    'Stafford, Ernest LSW': 'Licensed Social Worker',
    'Taylor, Cynthia RN': 'Registered Nurse',
    'Russo, Kelly RN': 'Registered Nurse',
    'Lightson, Ruth RN': 'Registered Nurse',
    'Brown, Lydia LPN': 'Practical Nurse',
    'Torres, Jennifer LPN': 'Practical Nurse',
    'Woods, Sharmone': 'Practical Nurse',
    'Oswald, Mark A MD': 'Doctor',
    'Ahsan, Shagufta MD': 'Doctor',
    'Childs, Brandis LCSW': 'Licensed Clinical Social Worker'
}

ph_valid_service_types = [
    'Chiropractor',
    'Dentist',
    'Dental Hygienist',
    'LCSW',
    'Nurse Midwife',
    'Nurse Practitioner',
    'OB/GYN',
    'Optometrist',
    'Physician',
    'Podiatry',
    'Psychologist,'
    'Unknown'
]

# this function is used to get the start date, end date, sheets, quarter_with_year, and months_with_year for a specified
# quarter and year.
def get_quarter_info(quarter, year, validate=True):

    quarter_dict = {                                                    # map quarters to months and sheet names
        'Q1': (['01', '02', '03'], ['Jan', 'Feb', 'Mar']),
        'Q2': (['04', '05', '06'], ['April', 'May', 'June']),
        'Q3': (['07', '08', '09'], ['July', 'Aug', 'Sept']),
        'Q4': (['10', '11', '12'], ['Oct', 'Nov', 'Dec'])
    }

    if validate and quarter not in quarter_dict:                        # Check if the quarter is valid
        print('Invalid quarter entered')
        return None, None, None

    months, sheet_months = quarter_dict[quarter]            # get months and sheet names for the quarter

    start_date = year + '-' + months[0] + '-01'             # get start and end dates for the quarter
    end_date = year + '-' + months[2] + '-30'

    # determine the number of schedules to create for reporting
    sheets = (
            ['Page 1', 'detail data'] +
            ['Support Schedule A - ' + month for month in sheet_months] +
            [ 'Support Schedule B - ' + month for month in sheet_months] +
            ['Support Schedule C - ' + month for month in sheet_months] +
            ['Support Schedule D - ' + month for month in sheet_months] +
            ['Support Schedule E - ' + month for month in sheet_months] +
            ['Support Schedule F - ' + month for month in sheet_months]
    )

    quarter_with_year = quarter + ' ' + year                                                                        # attach year to quarter
    months_with_year = [year + '-' + month for month in months]                             # attach year to specified months
    return start_date, end_date, sheets, quarter_with_year, months_with_year


def process_data_by_month(data_detail, months):
    # this function splits the dataset into 3 separate dataframes for each month. this is a requirement for the
    # NJDOH Wraparound Report.
    df1 = data_detail[(data_detail['service_month'] == months[0])].copy()
    df2 = data_detail[(data_detail['service_month'] == months[1])].copy()
    df3 = data_detail[(data_detail['service_month'] == months[2])].copy()

    print('\nunique claims for month 1:', df1['enc_nbr'].nunique())
    print('unique claims for month 2:', df2['enc_nbr'].nunique())
    print('unique claims for month 3:', df3['enc_nbr'].nunique())

    def sum_data(ds):
        ds = ds[['enc_nbr', 'service_type', 'hmo_name']].copy()
        ds['service_type'] = ds['service_type'].fillna('')

        ds = ds.groupby('enc_nbr').first().reset_index()
        ds = ds.groupby(['service_type', 'hmo_name']).agg({'enc_nbr': 'nunique'}).reset_index()
        ds = ds.pivot(index=['service_type'], columns='hmo_name', values='enc_nbr').reset_index()
        ds = ds.fillna(0)
        ds = ds[ds['service_type'] != ''].copy()

        # check if columns 001 to 005 exist in ds. if not add them
        for col in ['001', '002', '003', '004', '005']:
            if col not in ds.columns:
                ds[col] = 0

        ds = ds[['service_type', '001', '002', '003', '004', '005']].copy()  # arrange columns in order
        ds['total'] = ds['001'] + ds['002'] + ds['003'] + ds['004'] + ds['005']  # create totals column for 001 to 005
        return ds

    # process data for each month
    data1 = sum_data(df1)
    data2 = sum_data(df2)
    data3 = sum_data(df3)

    return data1, data2, data3
# NJDOH requires that the first row for an encounter have a encounter = 1 and every
# subsequent row for the same encounter have encounter = 0. the following code will
# update the encounter column to meet this requirement.
def update_encounter_column(dataset):
    last_enc_nbr = dataset['enc_nbr'][0]
    for row in dataset.itertuples():
        row_enc_nbr = dataset.at[row.Index, "enc_nbr"]
        if row.Index != 0:
            if row_enc_nbr == last_enc_nbr:
                dataset.at[row.Index, "encounter"] = 0
            else:
                dataset.at[row.Index, "encounter"] = 1
                last_enc_nbr = row_enc_nbr
    return dataset


def add_missing_service_types(column_value, d1, d2, d3):
    # this function adds the missing service types to the 3 dataframes used for the NJDOH Wraparound Report. The state
    # requirements that these columns are present
    # add columns to dataframes if they don't exist

    dict = {'service_type': [column_value], '001': [0], '002': [0], '003': [0], '004': [0], '005': [0], 'total': [0] }
    tempdf = pd.DataFrame(dict)

    # create a temp df instead of a dict
    if column_value not in d1['service_type'].values:
        d1 = pd.concat([d1, tempdf], ignore_index=True)
    if column_value not in d2['service_type'].values:
        d2 = pd.concat([d2, tempdf])
    if column_value not in d3['service_type'].values:
        d3 = pd.concat([d3, tempdf])
    return d1, d2, d3


def process_core_services(data1, data2, data3):
    """
    Add Core Services column and arrange column order for submission.

    Args:
        data1 (pd.DataFrame): The first dataframe.
        data2 (pd.DataFrame): The second dataframe.
        data3 (pd.DataFrame): The third dataframe.

    Returns:
        tuple: A tuple containing the updated dataframes.
    """
    def add_core_services_column(temp_df):
        temp_df['Core Services'] = 0  # create column Core Services and make it the first column

        # if service type is Physician set Core Services to 1
        temp_df.loc[temp_df['service_type'] == 'Physician', 'Core Services'] = 1
        temp_df.loc[temp_df['service_type'] == 'Nurse Practitioner', 'Core Services'] = 2
        temp_df.loc[temp_df['service_type'] == 'Dentist', 'Core Services'] = 3
        temp_df.loc[temp_df['service_type'] == 'Dental Hygienist', 'Core Services'] = 4
        temp_df.loc[temp_df['service_type'] == 'Nurse Midwife', 'Core Services'] = 5
        temp_df.loc[temp_df['service_type'] == 'OB/GYN', 'Core Services'] = 6
        temp_df.loc[temp_df['service_type'] == 'Podiatry', 'Core Services'] = 7
        temp_df.loc[temp_df['service_type'] == 'Chiropractor', 'Core Services'] = 8
        temp_df.loc[temp_df['service_type'] == 'Optometrist', 'Core Services'] = 9
        temp_df.loc[temp_df['service_type'] == 'LCSW', 'Core Services'] = 10
        temp_df.loc[temp_df['service_type'] == 'Psychologist', 'Core Services'] = 11
        temp_df = temp_df.sort_values(by=['Core Services'], ascending=True).reset_index(drop=True)
        return temp_df

    data1 = add_core_services_column(data1)
    data2 = add_core_services_column(data2)
    data3 = add_core_services_column(data3)

    # arrange column order for submission
    column_order = ['Core Services', 'service_type', '001', '002', '003', '004', '005', 'total']
    data1 = data1[column_order].copy()
    data2 = data2[column_order].copy()
    data3 = data3[column_order].copy()

    return data1, data2, data3
