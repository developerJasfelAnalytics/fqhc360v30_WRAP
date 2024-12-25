import pandas as pd

def process_schedule_B(schedB, months, horizon_capitation):
    """
    Process and Build Schedule B data for each month.

    Args:
        schedB (pd.DataFrame): The input data.
        months (list): List of months to process.
        horizon_capitation (float): The capitation amount for Horizon.

    Returns:
        tuple: A tuple containing the processed dataframes for each month.
    """
    month_names = [pd.to_datetime(month).strftime('%b') for month in months]

    schedB1 = schedB[schedB['service_month'] == months[0]].copy()
    schedB2 = schedB[schedB['service_month'] == months[1]].copy()
    schedB3 = schedB[schedB['service_month'] == months[2]].copy()

    def enhance_schedB(dfsched):
        dfsched['date_of_service'] = pd.to_datetime(dfsched['date_of_service'])
        dfsched['claim_payment_amount'] = dfsched['claim_payment_amount'].abs()

        df_schedB = dfsched.groupby(['service_month', 'hmo_name']).agg({'claim_payment_amount': 'sum'}).reset_index()
        df_schedB = df_schedB.pivot(index=['service_month'], columns='hmo_name', values='claim_payment_amount')
        df_schedB = df_schedB.reset_index()
        df_schedB['Num'] = 3

        for col in ['001', '002', '003', '004', '005']:
            if col not in df_schedB.columns:
                df_schedB[col] = 0

        df_schedB = df_schedB[['Num', 'service_month', '001', '002', '003', '004', '005']]
        df_schedB.iloc[0, 1] = 'Fee for Service'

        df_schedB_data = [
            [2, 'Capitation Receipts', 0, 0, horizon_capitation, 0, 0],
            [4, 'TLP Receipts', 0, 0, 0, 0, 0],
            [5, 'Other (Specify)', 0, 0, 0, 0, 0],
            [6, 'Other (Specify)', 0, 0, 0, 0, 0],
        ]

        df_schedB_data = pd.DataFrame(df_schedB_data, columns=['Num', 'service_month', '001', '002', '003', '004', '005'])
        df_schedB = pd.concat([df_schedB_data, df_schedB], ignore_index=True)
        df_schedB = df_schedB.sort_values('Num', ascending=True).reset_index(drop=True)
        return df_schedB

    schedB1_sum = enhance_schedB(schedB1)
    schedB2_sum = enhance_schedB(schedB2)
    schedB3_sum = enhance_schedB(schedB3)

    return schedB1_sum, schedB2_sum, schedB3_sum


def write_scheduleA(health_center, fqhc_billing_number, ph_hmo_dict, workbook, worksheet, i, j, year):
    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True, 'align': 'center'})
    bold_left = workbook.add_format({'bold': True, 'align': 'left'})

    # add a red format to use to highlight cells
    red_format = workbook.add_format({'bold': True, 'font_color': '#9C0006', 'align': 'center'})

    # add titles to worksheet
    worksheet.write('B1', 'Federally Qualified health Center Name: ' + health_center + ' ', bold_left)
    worksheet.write('B2', 'MEDICAID MANAGED CARE ENCOUNTER DETAIL', bold_left)
    worksheet.write('C3', 'Reporting Month: ' + j + ' ' + year, bold_left)
    worksheet.write('F2', 'FQHC Number: ' + fqhc_billing_number, bold_left)
    worksheet.write('H1', 'Worksheet 2', bold)
    worksheet.write('H2', 'Support Schedule A', bold)
    worksheet.write('H4', 'Total', bold)
    worksheet.write('H5', 'Medicald', bold)
    worksheet.write('H6', 'HMO', bold)
    worksheet.write('H7', 'Encounters', bold)

    worksheet.write('C6', 'HMO 001', bold)
    worksheet.write('D6', 'HMO 002', bold)
    worksheet.write('E6', 'HMO 003', bold)
    worksheet.write('F6', 'HMO 004', bold)
    worksheet.write('G6', 'HMO 005', bold)

    worksheet.write('C7', list(ph_hmo_dict.keys())[0], red_format)
    worksheet.write('D7', list(ph_hmo_dict.keys())[1], red_format)
    worksheet.write('E7', list(ph_hmo_dict.keys())[2], red_format)
    worksheet.write('F7', list(ph_hmo_dict.keys())[3], red_format)
    worksheet.write('G7', list(ph_hmo_dict.keys())[4], red_format)

    worksheet.write('C8', '(1)', bold)
    worksheet.write('D8', '(2)', bold)
    worksheet.write('E8', '(3)', bold)
    worksheet.write('F8', '(4)', bold)
    worksheet.write('G8', '(5)', bold)
    worksheet.write('H8', '(7)', bold)

    # set column width
    worksheet.set_column(0, 0, 12)
    worksheet.set_column(1, 1, 30)
    worksheet.set_column(2, 7, 20)

    worksheet.write('A9', 'Core Services', bold)

    # create a totals row for each column
    worksheet.write('B22', 'Total Payable Encounters')
    worksheet.write_formula('C22', '=SUM(C10:C' + str(len(i.index) + 9) + ')')
    worksheet.write_formula('D22', '=SUM(D10:D' + str(len(i.index) + 9) + ')')
    worksheet.write_formula('E22', '=SUM(E10:E' + str(len(i.index) + 9) + ')')
    worksheet.write_formula('F22', '=SUM(F10:F' + str(len(i.index) + 9) + ')')
    worksheet.write_formula('G22', '=SUM(G10:G' + str(len(i.index) + 9) + ')')
    worksheet.write_formula('H22', '=SUM(H10:H' + str(len(i.index) + 9) + ')')


def write_scheduleB(health_center, fqhc_billing_number, ph_hmo_dict, workbook, worksheet, i, j, year):
    # document this function

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True, 'align': 'center'})
    bold_left = workbook.add_format({'bold': True, 'align': 'left'})

    # add a red format to use to highlight cells
    red_format = workbook.add_format({'bold': True, 'font_color': '#9C0006', 'align': 'center'})

    # add titles to worksheet
    worksheet.write('B1', 'Federally Qualified health Center Name: ' + health_center, bold_left)
    worksheet.write('B2', 'MEDICAID MANAGED CARE ENCOUNTER DETAIL', bold_left)
    worksheet.write('C3', 'Reporting Month: ' + j + ' ' + year, bold_left)
    worksheet.write('F2', 'FQHC Number: ' + fqhc_billing_number, bold_left)
    worksheet.write('H1', 'Worksheet 2', bold)
    worksheet.write('H2', 'Support Schedule A', bold)
    worksheet.write('H4', 'Total', bold)
    worksheet.write('H5', 'Medicald', bold)
    worksheet.write('H6', 'HMO', bold)
    worksheet.write('H7', 'Encounters', bold)

    worksheet.write('C6', 'HMO 001', bold)
    worksheet.write('D6', 'HMO 002', bold)
    worksheet.write('E6', 'HMO 003', bold)
    worksheet.write('F6', 'HMO 004', bold)
    worksheet.write('G6', 'HMO 005', bold)

    worksheet.write('B7', 'HMO Name', bold)
    worksheet.write('C7', list(ph_hmo_dict.keys())[0], red_format)
    worksheet.write('D7', list(ph_hmo_dict.keys())[1], red_format)
    worksheet.write('E7', list(ph_hmo_dict.keys())[2], red_format)
    worksheet.write('F7', list(ph_hmo_dict.keys())[3], red_format)
    worksheet.write('G7', list(ph_hmo_dict.keys())[4], red_format)

    worksheet.write_formula('H08', '=SUM(C8 :G8)')  # add total encounters
    worksheet.write_formula('H09', '=SUM(C9 :G9)')
    worksheet.write_formula('H10', '=SUM(C10:G10)')
    worksheet.write_formula('H11', '=SUM(C11:G11)')
    worksheet.write_formula('H12', '=SUM(C12:G12)')
    worksheet.write_formula('H13', '=SUM(C13:G13)')

    worksheet.write('B13', 'Total Receipts')  # add total receipts
    worksheet.write_formula('C13', '=SUM(C08:C' + str(len(i.index) + 7) + ')')
    worksheet.write_formula('D13', '=SUM(D08:D' + str(len(i.index) + 7) + ')')
    worksheet.write_formula('E13', '=SUM(E08:E' + str(len(i.index) + 7) + ')')
    worksheet.write_formula('F13', '=SUM(F08:F' + str(len(i.index) + 7) + ')')
    worksheet.write_formula('G13', '=SUM(G08:G' + str(len(i.index) + 7) + ')')
    worksheet.write_formula('H13', '=SUM(H08:H' + str(len(i.index) + 7) + ')')

    worksheet.set_column(0, 0, 12)  # set column width
    worksheet.set_column(1, 1, 30)
    worksheet.set_column(2, 7, 20)



def build_spreadsheet_and_schedules(health_center, fqhc_billing_number, ph_hmo_dict, year, workdir, quarter_with_year, report_type,
                                    page1_data, data_detail,
                                    data1, data2, data3,
                                    schedB1_sum, schedB2_sum, schedB3_sum, schedC_data,
                                    sheets, month_names):
    # ----------------------------------------------------
    # build Schedules
    # ----------------------------------------------------
    # specify where to create the spreadsheet
    file_name = (workdir + 'wrap_to_submit\\ ' + quarter_with_year + ' ' +
                 report_type + ' - DRAFT.xlsx')
    print('creating file: ', file_name)

    sheet_schedA = 'Support Schedule A - '
    sheet_schedB = 'Support Schedule B -'
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')

    # create master file for submission
    with pd.ExcelWriter(file_name) as writer:
        page1_data.to_excel(writer, sheet_name='Page 1', startrow=9, header=False, index=False)
        data_detail.to_excel(writer, sheet_name='data detail', index=False)
        data1.to_excel(writer, sheet_name=sheet_schedA + ' ' + month_names[0], startrow=9, header=False, index=False)
        data2.to_excel(writer, sheet_name=sheet_schedA + ' ' + month_names[1], startrow=9, header=False, index=False)
        data3.to_excel(writer, sheet_name=sheet_schedA + ' ' + month_names[2], startrow=9, header=False, index=False)
        schedB1_sum.to_excel(writer, sheet_name=sheet_schedB + ' ' + month_names[0], startrow=7, header=False,
                             index=False)
        schedB2_sum.to_excel(writer, sheet_name=sheet_schedB + ' ' + month_names[1], startrow=7, header=False,
                             index=False)
        schedB3_sum.to_excel(writer, sheet_name=sheet_schedB + ' ' + month_names[2], startrow=7, header=False,
                             index=False)
        schedC_data.to_excel(writer, sheet_name='Support Schedule C ' + month_names[0], startrow=10, header=False,
                             index=False)
        schedC_data.to_excel(writer, sheet_name='Support Schedule C ' + month_names[1], startrow=10, header=False,
                             index=False)
        schedC_data.to_excel(writer, sheet_name='Support Schedule C ' + month_names[2], startrow=10, header=False,
                             index=False)
        schedC_data.to_excel(writer, sheet_name='Support Schedule D ' + month_names[0], startrow=10, header=False,
                             index=False)
        schedC_data.to_excel(writer, sheet_name='Support Schedule D ' + month_names[1], startrow=10, header=False,
                             index=False)
        schedC_data.to_excel(writer, sheet_name='Support Schedule D ' + month_names[2], startrow=10, header=False,
                             index=False)
        schedC_data.to_excel(writer, sheet_name='Support Schedule E ' + month_names[0], startrow=10, header=False,
                             index=False)
        schedC_data.to_excel(writer, sheet_name='Support Schedule E ' + month_names[1], startrow=10, header=False,
                             index=False)
        schedC_data.to_excel(writer, sheet_name='Support Schedule E ' + month_names[2], startrow=10, header=False,
                             index=False)
        schedC_data.to_excel(writer, sheet_name='Support Schedule F ' + month_names[0], startrow=10, header=False,
                             index=False)
        schedC_data.to_excel(writer, sheet_name='Support Schedule F ' + month_names[1], startrow=10, header=False,
                             index=False)
        schedC_data.to_excel(writer, sheet_name='Support Schedule F ' + month_names[2], startrow=10, header=False,
                             index=False)

    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    workbook = writer.book

    sheet_data = [page1_data, data_detail,
                  data1, data2, data3,
                  schedB1_sum, schedB2_sum, schedB3_sum,
                  schedC_data, schedC_data, schedC_data,
                  schedC_data, schedC_data, schedC_data,  # this is used for Schedule D
                  schedC_data, schedC_data, schedC_data,  # this is used for Schedule E
                  schedC_data, schedC_data, schedC_data  # this is used for Schedule F
                  ]  # used for Maria reviewed

    for i, j in zip(sheet_data, sheets):
        print('writing ', j, ' to excel file...')

        # Add a bold format to use to highlight cells.
        bold = workbook.add_format({'bold': True, 'align': 'center'})
        bold_left = workbook.add_format({'bold': True, 'align': 'left'})

        # add a red format to use to highlight cells
        red_format = workbook.add_format({'bold': True, 'font_color': '#9C0006', 'align': 'center'})
        red_format_left = workbook.add_format({'bold': True, 'font_color': '#9C0006', 'align': 'left'})

        if j == 'Page 1':
            print('writing Page 1 to excel file...')
            i.to_excel(writer, sheet_name=j, startrow=9, header=False, index=False)
            worksheet = writer.sheets[j]

            def write_page1(worksheet, bold_left, bold, red_format_left, quarter_with_year):
                # add titles to worksheet
                worksheet.write('B1', 'FQHC WRAPAROUND RECONCILIATION REPORT', bold_left)
                worksheet.write('A3', 'Prov No.', bold_left)
                worksheet.write('B3', fqhc_billing_number, bold_left)
                worksheet.write('A5', 'Prov Name', bold_left)
                worksheet.write('B5', health_center, bold_left)
                worksheet.write('C6', 'Initial &', bold)
                worksheet.write('B7', 'CALENDAR YEAR QUARTER:', bold_left)
                worksheet.write('C7', 'Revision to Initial', bold)
                worksheet.write('D7', 'Reconciliation', bold)
                worksheet.write('B8', quarter_with_year, bold_left)
                worksheet.write('C8', 'Payment', bold)
                worksheet.write('D8', quarter_with_year, bold)
                worksheet.write('A10', 'A', bold)
                worksheet.write('B10', 'Medicaid Managed Care Encounter Approved', bold_left)
                worksheet.write('A11', 'B', bold)
                worksheet.write('B11', 'Medicaid PPS (pps alternative methodology rate)', bold_left)
                worksheet.write('A12', 'C', bold)
                worksheet.write('B12', 'Total Payment  (A times B)', bold_left)
                worksheet.write('A13', 'D', bold)
                worksheet.write('B13', 'Medicaid Managed Care Receipts', bold_left)
                worksheet.write('A14', 'E', bold)
                worksheet.write('B14', 'Excluded Vaccination Receipts', red_format_left)
                worksheet.write('A15', 'F', bold)
                worksheet.write('B15', 'Difference    (C less D) + E', bold_left)
                worksheet.write('A16', 'G', bold)
                worksheet.write('B16', 'Amount Due/ (From)                  (column E less column D)', bold_left)

                # set column width
                worksheet.set_column(0, 0, 8)
                worksheet.set_column(1, 1, 60)
                worksheet.set_column(2, 3, 10)

            # Call the function with the required parameters
            write_page1(worksheet, bold_left, bold, red_format_left, quarter_with_year)

        elif j == 'detail data':
            i.to_excel(writer, sheet_name=j, index=False)
            worksheet = writer.sheets[j]

            # set column width
            worksheet.set_column(0, 2, 20)
            worksheet.set_column(3, 14, 30)

        elif 'Schedule A' in j:
            i.to_excel(writer, sheet_name=j, startrow=9, header=False, index=False)
            worksheet = writer.sheets[j]

            write_scheduleA(health_center, fqhc_billing_number, ph_hmo_dict, workbook, worksheet, i, j, year)

        elif 'Schedule B' in j:
            i.to_excel(writer, sheet_name=j, startrow=7, header=False, index=False)  # write data to excel file
            worksheet = writer.sheets[j]

            write_scheduleB(health_center, fqhc_billing_number, ph_hmo_dict, workbook, worksheet, i, j, year)

        elif 'Schedule C' in j:
            # Add a bold format to use to highlight cells.
            bold = workbook.add_format(
                {'bold': True, 'align': 'center'})  # Add a bold format to use to highlight cells.
            bold_left = workbook.add_format({'bold': True, 'align': 'left'})

            i.to_excel(writer, sheet_name=j, startrow=7, header=False, index=False)
            worksheet = writer.sheets[j]

            worksheet.write('B1', 'Federally Qualified health Center Name: ' + health_center, bold_left)
            worksheet.write('B2', 'Medicaid Managed Care Delivery Encounters Detail', bold_left)

            worksheet.write('F1', 'FQHC Number: ' + fqhc_billing_number, bold_left)
            worksheet.write('H1', 'Worksheet 2', bold)
            worksheet.write('H2', 'Support Schedule C', bold)
            worksheet.write('B4', 'A', bold)
            worksheet.write('C4', 'B', bold)
            worksheet.write('D4', 'C', bold)
            worksheet.write('E4', 'D', bold)
            worksheet.write('F4', 'E', bold)
            worksheet.write('G4', 'F', bold)
            worksheet.write('H4', 'G', bold)

            worksheet.write('C5', 'HMO 001', bold)
            worksheet.write('D5', 'HMO 002', bold)
            worksheet.write('E5', 'HMO 003', bold)
            worksheet.write('F5', 'HMO 004', bold)
            worksheet.write('G5', 'HMO 005', bold)
            worksheet.write('H5', 'Total Medicaid', bold)

            worksheet.write('B6', 'HMO Name', bold_left)
            worksheet.write('C6', list(ph_hmo_dict.keys())[0], red_format)
            worksheet.write('D6', list(ph_hmo_dict.keys())[1], red_format)
            worksheet.write('E6', list(ph_hmo_dict.keys())[2], red_format)
            worksheet.write('F6', list(ph_hmo_dict.keys())[3], red_format)
            worksheet.write('G6', list(ph_hmo_dict.keys())[4], red_format)
            worksheet.write('H6', 'Delivery', bold)

            worksheet.write('B8', 'Delivery Procedure Code', bold_left)
            worksheet.write('H8', 'Encounters', bold)

            worksheet.write('A36', 31)
            worksheet.write('B36', 'Total (Lines 8-34)', bold_left)
            worksheet.write_formula('C36', '=SUM(C8:C33)')
            worksheet.write_formula('D36', '=SUM(D8:D33)')
            worksheet.write_formula('E36', '=SUM(E8:E33)')
            worksheet.write_formula('F36', '=SUM(F8:F33)')
            worksheet.write_formula('G36', '=SUM(G8:G33)')
            worksheet.write_formula('H36', '=SUM(H8:H33)')

            # write data to excel file. write  in columna A1 to A3 1 to 3
            worksheet.write('A1', 1)
            worksheet.write('A2', 2)
            worksheet.write('A3', 3)
            worksheet.write('A8', '4')

            # Write values in column 'A' from row 6 to 36
            for row, value in enumerate(range(5, 31), start=9):
                worksheet.write(f'A{row}', value)

            worksheet.set_column(0, 0, 5)  # set column width
            worksheet.set_column(1, 1, 30)
            worksheet.set_column(2, 7, 20)

        elif 'Schedule D' in j:
            # Add a bold format to use to highlight cells.
            bold = workbook.add_format(
                {'bold': True, 'align': 'center'})  # Add a bold format to use to highlight cells.
            bold_left = workbook.add_format({'bold': True, 'align': 'left'})

            i.to_excel(writer, sheet_name=j, startrow=7, header=False, index=False)
            worksheet = writer.sheets[j]

            worksheet.write('B1', 'Federally Qualified health Center Name: ' + health_center, bold_left)
            worksheet.write('B2', 'Medicaid Managed Care Delivery Receipts', bold_left)

            worksheet.write('F1', 'FQHC Number: ' + fqhc_billing_number, bold_left)
            worksheet.write('H1', 'Worksheet 2', bold)
            worksheet.write('H2', 'Support Schedule D', bold)
            worksheet.write('B4', 'A', bold)
            worksheet.write('C4', 'B', bold)
            worksheet.write('D4', 'C', bold)
            worksheet.write('E4', 'D', bold)
            worksheet.write('F4', 'E', bold)
            worksheet.write('G4', 'F', bold)
            worksheet.write('H4', 'G', bold)

            worksheet.write('C5', 'HMO 001', bold)
            worksheet.write('D5', 'HMO 002', bold)
            worksheet.write('E5', 'HMO 003', bold)
            worksheet.write('F5', 'HMO 004', bold)
            worksheet.write('G5', 'HMO 005', bold)
            worksheet.write('H5', 'Total Medicaid', bold)

            worksheet.write('B6', 'HMO Name', bold_left)
            worksheet.write('C6', list(ph_hmo_dict.keys())[0], red_format)
            worksheet.write('D6', list(ph_hmo_dict.keys())[1], red_format)
            worksheet.write('E6', list(ph_hmo_dict.keys())[2], red_format)
            worksheet.write('F6', list(ph_hmo_dict.keys())[3], red_format)
            worksheet.write('G6', list(ph_hmo_dict.keys())[4], red_format)
            worksheet.write('H6', 'Delivery', bold)

            worksheet.write('B8', 'Delivery Procedure Code', bold_left)
            worksheet.write('H8', 'Receipts', bold)

            worksheet.write('A36', 31)
            worksheet.write('B36', 'Total (Lines 8-34)', bold_left)
            worksheet.write_formula('C36', '=SUM(C8:C33)')
            worksheet.write_formula('D36', '=SUM(D8:D33)')
            worksheet.write_formula('E36', '=SUM(E8:E33)')
            worksheet.write_formula('F36', '=SUM(F8:F33)')
            worksheet.write_formula('G36', '=SUM(G8:G33)')
            worksheet.write_formula('H36', '=SUM(H8:H33)')

            # write data to excel file. write  in columna A1 to A3 1 to 3
            worksheet.write('A1', 1)
            worksheet.write('A2', 2)
            worksheet.write('A3', 3)
            worksheet.write('A8', '4')

            # Write values in column 'A' from row 6 to 36
            for row, value in enumerate(range(5, 31), start=9):
                worksheet.write(f'A{row}', value)

            worksheet.set_column(0, 0, 5)  # set column width
            worksheet.set_column(1, 1, 30)
            worksheet.set_column(2, 7, 20)

        elif 'Schedule E' in j:
            # Add a bold format to use to highlight cells.
            bold = workbook.add_format(
                {'bold': True, 'align': 'center'})  # Add a bold format to use to highlight cells.
            bold_left = workbook.add_format({'bold': True, 'align': 'left'})

            i.to_excel(writer, sheet_name=j, startrow=7, header=False, index=False)
            worksheet = writer.sheets[j]

            worksheet.write('B1', 'Federally Qualified health Center Name: ' + health_center, bold_left)
            worksheet.write('B2', 'Medicaid Managed Care OB/GYN Surgical Encounters Detail', bold_left)

            worksheet.write('F1', 'FQHC Number: ' + fqhc_billing_number, bold_left)
            worksheet.write('H1', 'Worksheet 2', bold)
            worksheet.write('H2', 'Support Schedule D', bold)
            worksheet.write('B4', 'A', bold)
            worksheet.write('C4', 'B', bold)
            worksheet.write('D4', 'C', bold)
            worksheet.write('E4', 'D', bold)
            worksheet.write('F4', 'E', bold)
            worksheet.write('G4', 'F', bold)
            worksheet.write('H4', 'G', bold)

            worksheet.write('C5', 'HMO 001', bold)
            worksheet.write('D5', 'HMO 002', bold)
            worksheet.write('E5', 'HMO 003', bold)
            worksheet.write('F5', 'HMO 004', bold)
            worksheet.write('G5', 'HMO 005', bold)
            worksheet.write('H5', 'Total Medicaid', bold)

            worksheet.write('B6', 'HMO Name', bold_left)
            worksheet.write('C6', list(ph_hmo_dict.keys())[0], red_format)
            worksheet.write('D6', list(ph_hmo_dict.keys())[1], red_format)
            worksheet.write('E6', list(ph_hmo_dict.keys())[2], red_format)
            worksheet.write('F6', list(ph_hmo_dict.keys())[3], red_format)
            worksheet.write('G6', list(ph_hmo_dict.keys())[4], red_format)
            worksheet.write('H6', 'OB/GYN', bold)

            worksheet.write('B8', 'OB/GYN Surgical Delivery Procedure Code', bold_left)
            worksheet.write('H8', 'Surgical Encounters', bold)

            worksheet.write('A36', 31)
            worksheet.write('B36', 'Total (Lines 8-34)', bold_left)
            worksheet.write_formula('C36', '=SUM(C8:C33)')
            worksheet.write_formula('D36', '=SUM(D8:D33)')
            worksheet.write_formula('E36', '=SUM(E8:E33)')
            worksheet.write_formula('F36', '=SUM(F8:F33)')
            worksheet.write_formula('G36', '=SUM(G8:G33)')
            worksheet.write_formula('H36', '=SUM(H8:H33)')

            # write data to excel file. write  in columna A1 to A3 1 to 3
            worksheet.write('A1', 1)
            worksheet.write('A2', 2)
            worksheet.write('A3', 3)
            worksheet.write('A6', 4)

            # Write values in column 'A' from row 6 to 36
            for row, value in enumerate(range(5, 31), start=9):
                worksheet.write(f'A{row}', value)

            worksheet.set_column(0, 0, 5)  # set column width
            worksheet.set_column(1, 1, 30)
            worksheet.set_column(2, 7, 20)

        elif 'Schedule F' in j:
            # Add a bold format to use to highlight cells.
            bold = workbook.add_format(
                {'bold': True, 'align': 'center'})  # Add a bold format to use to highlight cells.
            bold_left = workbook.add_format({'bold': True, 'align': 'left'})

            i.to_excel(writer, sheet_name=j, startrow=7, header=False, index=False)
            worksheet = writer.sheets[j]

            worksheet.write('B1', 'Federally Qualified health Center Name: ' + health_center, bold_left)
            worksheet.write('B2', 'Medicaid Managed Care OB/GYN Surgical Encounters Detail', bold_left)

            worksheet.write('F1', 'FQHC Number: ' + fqhc_billing_number, bold_left)
            worksheet.write('H1', 'Worksheet 2', bold)
            worksheet.write('H2', 'Support Schedule D', bold)
            worksheet.write('B4', 'A', bold)
            worksheet.write('C4', 'B', bold)
            worksheet.write('D4', 'C', bold)
            worksheet.write('E4', 'D', bold)
            worksheet.write('F4', 'E', bold)
            worksheet.write('G4', 'F', bold)
            worksheet.write('H4', 'G', bold)

            worksheet.write('C5', 'HMO 001', bold)
            worksheet.write('D5', 'HMO 002', bold)
            worksheet.write('E5', 'HMO 003', bold)
            worksheet.write('F5', 'HMO 004', bold)
            worksheet.write('G5', 'HMO 005', bold)
            worksheet.write('H5', 'Total Medicaid', bold)

            worksheet.write('B6', 'HMO Name', bold_left)
            worksheet.write('C6', list(ph_hmo_dict.keys())[0], red_format)
            worksheet.write('D6', list(ph_hmo_dict.keys())[1], red_format)
            worksheet.write('E6', list(ph_hmo_dict.keys())[2], red_format)
            worksheet.write('F6', list(ph_hmo_dict.keys())[3], red_format)
            worksheet.write('G6', list(ph_hmo_dict.keys())[4], red_format)
            worksheet.write('H6', 'OB/GYN', bold)

            worksheet.write('B8', 'OB/GYN Surgical Delivery Procedure Code', bold_left)
            worksheet.write('H8', 'Surgical Encounters', bold)

            worksheet.write('A36', 31)
            worksheet.write('B36', 'Total (Lines 8-34)', bold_left)
            worksheet.write_formula('C36', '=SUM(C8:C33)')
            worksheet.write_formula('D36', '=SUM(D8:D33)')
            worksheet.write_formula('E36', '=SUM(E8:E33)')
            worksheet.write_formula('F36', '=SUM(F8:F33)')
            worksheet.write_formula('G36', '=SUM(G8:G33)')
            worksheet.write_formula('H36', '=SUM(H8:H33)')

            # write data to Excel file. write  in columna A1 to A3 1 to 3
            worksheet.write('A1', 1)
            worksheet.write('A2', 2)
            worksheet.write('A3', 3)
            worksheet.write('A8', '4')

            # Write values in column 'A' from row 6 to 36
            for row, value in enumerate(range(5, 31), start=9):
                worksheet.write(f'A{row}', value)

            worksheet.set_column(0, 0, 5)  # set column width
            worksheet.set_column(1, 1, 30)
            worksheet.set_column(2, 7, 20)

    writer.close()
