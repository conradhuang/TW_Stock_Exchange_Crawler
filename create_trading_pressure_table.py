﻿# coding=utf-8
import xlsxwriter
import glob
import csv
import os
import re
import string

from os import mkdir
from os.path import isdir
from xlsxwriter.utility import xl_rowcol_to_cell

WorkingDirectory = os.getcwd()
TSE_RAW_DATA_FOLDER = 'tse_trading_raw_data'
OTC_RAW_DATA_FOLDER = 'otc_trading_raw_data'
RAW_DATA_FOLDER = [TSE_RAW_DATA_FOLDER, OTC_RAW_DATA_FOLDER]
FOLDER_SAVE = 'trading_pressure_data'
FILE_NAME = 'trading_pressure_table.xlsm'
BUY_VALUE = 0.6
SELL_VALUE = 1.4

titles = ['時間', '成交量', '開盤', '最高', '最低', '收盤', '收盤價差', '', '買盤1', '賣盤1', '買盤2', '賣盤2', '',
          '漲幅總和', '跌幅總和', '', '買盤力道張數', '賣盤力道張數', '', '20日總買盤', '20日總賣盤', '', '買賣壓力道比例',
          '買點', '賣點']
row_end = 248
row_init = 0
col_init = 0

summarybook = xlsxwriter.Workbook('./{}'.format(FILE_NAME))
summary_format = summarybook.add_format()
summary_format.set_font_color('red')
summary_format.set_align('center')
summary_date_format = summarybook.add_format()
summary_date_format.set_align('center')
# Add the VBA project binary.
summarybook.add_vba_project('./excel_vba_tools/UpdateTableValues.bin')

format = summarybook.add_format()
format.set_border(1)
format.set_bg_color('#cccccc')
format.set_align('center')

link_format = summarybook.add_format()
link_format.set_font_color('blue')
# format.set_bold()


def formula(worksheet):
    row_init = 0
    col_init = 0
    row_end = 248
    for row in range(1, row_end, 1):
        # current row value
        col_B = xl_rowcol_to_cell(row + 1, 1)
        col_C = xl_rowcol_to_cell(row + 1, 2)
        col_D = xl_rowcol_to_cell(row + 1, 3)
        col_E = xl_rowcol_to_cell(row + 1, 4)
        col_F_before = xl_rowcol_to_cell(row, 5)
        col_F = xl_rowcol_to_cell(row + 1, 5)
        col_G = xl_rowcol_to_cell(row + 1, 6)
        col_I = xl_rowcol_to_cell(row + 1, 8)
        col_J = xl_rowcol_to_cell(row + 1, 9)
        col_K = xl_rowcol_to_cell(row + 1, 10)
        col_L = xl_rowcol_to_cell(row + 1, 11)
        col_N = xl_rowcol_to_cell(row + 1, 13)
        col_O = xl_rowcol_to_cell(row + 1, 14)
        col_Q = xl_rowcol_to_cell(row + 1, 16)
        col_R = xl_rowcol_to_cell(row + 1, 17)
        col_T = xl_rowcol_to_cell(row + 1, 19)
        col_U = xl_rowcol_to_cell(row + 1, 20)
        if row > 19:
            col_Q_before = xl_rowcol_to_cell(row - 18, 16)
            col_R_before = xl_rowcol_to_cell(row - 18, 17)
        # 收盤價差
        first_row = xl_rowcol_to_cell(row, 5)
        next_row = xl_rowcol_to_cell(row + 1, 5)
        worksheet.write_formula(row + 1, col_init + 6, "%s-%s" % (next_row, first_row))
        # 買盤1
        # =ABS(IF(G5>0,C5-F4,D5-C5))
        worksheet.write_formula(row + 1, col_init + 8,
                                "ABS(IF(%s>0,%s-%s,%s-%s))" % (col_G, col_C, col_F_before, col_D, col_C))
        # 賣盤1
        # =ABS(IF(G5>0,C5-E5,F4-C5))
        worksheet.write_formula(row + 1, col_init + 9,
                                "ABS(IF(%s>0,%s-%s,%s-%s))" % (col_G, col_C, col_E, col_F_before, col_C))
        # 買盤2
        # =ABS(IF(G5>0,D5-E5,F5-E5))
        worksheet.write_formula(row + 1, col_init + 10,
                                "ABS(IF(%s>0,%s-%s,%s-%s))" % (col_G, col_D, col_E, col_F, col_E))
        # 賣盤2
        # =ABS(IF(G5>0,D5-F5,D5-E5))
        worksheet.write_formula(row + 1, col_init + 11,
                                "ABS(IF(%s>0,%s-%s,%s-%s))" % (col_G, col_D, col_F, col_D, col_E))
        # 漲幅總和
        # =I5+K5 (if I to J are equal to 0, write 0.0001)
        # =IF(Ix=0,IF(Jx=0, IF(Kx=0, IF(Lx=0, 0.0001, Ix+Kx), Ix+Kx), Ix+Kx), Ix+Kx)
        worksheet.write_formula(row + 1, col_init + 13,
                                "IF(%s=0,IF(%s=0, IF(%s=0, if(%s=0, 0.0001, %s+%s), %s+%s), %s+%s), %s+%s)" % (
                                col_I, col_J, col_K, col_L, col_I, col_K, col_I, col_K, col_I, col_K, col_I, col_K))
        # 跌幅總和
        # =J5+L5 (if I to J are equal to 0, write 0.0001)
        # =IF(Ix=0,IF(Jx=0, IF(Kx=0, IF(Lx=0, 0.0001, Jx+Lx), Jx+Lx), Jx+Lx), Jx+Lx)
        worksheet.write_formula(row + 1, col_init + 14,
                                "IF(%s=0,IF(%s=0, IF(%s=0, if(%s=0, 0.0001, %s+%s), %s+%s), %s+%s), %s+%s)" % (
                                col_I, col_J, col_K, col_L, col_J, col_L, col_J, col_L, col_J, col_L, col_J, col_L))
        # 買盤力道張數
        # =INT(B5*(N5/(N5+O5)))
        worksheet.write_formula(row + 1, col_init + 16, "INT(%s*(%s/(%s+%s)))" % (col_B, col_N, col_N, col_O))
        # 賣盤力道張數
        # =INT(B5*(O5/(N5+O5)))
        worksheet.write_formula(row + 1, col_init + 17, "INT(%s*(%s/(%s+%s)))" % (col_B, col_O, col_N, col_O))

        if row > 19:
            # 20天總買盤
            # =SUM(Q5:Q24)
            worksheet.write_formula(row + 1, col_init + 19, "SUM(%s:%s)" % (col_Q_before, col_Q))
            # 20天總賣盤
            # =SUM(R5:R24)
            worksheet.write_formula(row + 1, col_init + 20, "SUM(%s:%s)" % (col_R_before, col_R))
            # 買賣壓力道比例
            # =U24/T24
            worksheet.write_formula(row + 1, col_init + 22, "%s/%s" % (col_U, col_T))
            # 買,賣 點臨介值
            worksheet.write(row + 1, col_init + 23, BUY_VALUE)
            worksheet.write(row + 1, col_init + 24, SELL_VALUE)


def chart_def(spreadbook, worksheet, f_short_name):
    chart = spreadbook.add_chart({'type': 'stock'})
    # Add a series for each of the High-Low-Close columns.
    chart.add_series({
        'categories': '=%s!$A$22:$A$249' % (f_short_name),
        'values': '=%s!$D$22:$D$249' % (f_short_name),
    })

    chart.add_series({
        'categories': '=%s!$A$22:$A$249' % (f_short_name),
        'values': '=%s!$E$22:$E$249' % (f_short_name),
    })

    chart.add_series({
        'categories': '=%s!$A$22:$A$249' % (f_short_name),
        'values': '=%s!$F$22:$F$249' % (f_short_name),
    })

    chart.set_title({'name': '買賣壓力道比值'})
    chart.set_x_axis({'name': 'Date'})
    chart.set_y_axis({'name': 'Price'})
    chart.set_size({'width': 720, 'height': 360})
    chart.set_legend({'none': True})

    chart.set_up_down_bars({
        'up': {
            'fill': {'color': '#00B050'},
            'border': {'color': 'black'}
        },
        'down': {
            'fill': {'color': 'red'},
            'border': {'color': 'black'},
        },
    })

    # Create threshold-line as the secondary chart.
    line_chart = spreadbook.add_chart({'type': 'line'})

    line_chart.add_series({
        'categories': '=%s!$A$22:$A$249' % (f_short_name),
        'values': '=%s!$W$22:$W$249' % (f_short_name),
        'line': {'color': 'gray'},
        'y2_axis': True,
    })

    line_chart.add_series({
        'categories': '=%s!$A$22:$A$249' % (f_short_name),
        'values': '=%s!$X$22:$X$249' % (f_short_name),
        'line': {'color': 'red'},
        'y2_axis': True,
    })

    line_chart.add_series({
        'categories': '=%s!$A$22:$A$249' % (f_short_name),
        'values': '=%s!$Y$22:$Y$249' % (f_short_name),
        'line': {'color': 'green'},
        'y2_axis': True,
    })
    line_chart.set_size({'width': 720, 'height': 360})

    chart.combine(line_chart)
    worksheet.insert_chart('E251', chart)


def get_last_rows_of_CSV(f_name):
    with open('{}/{}'.format(TSE_RAW_DATA_FOLDER, f_name), 'r', encoding='utf-8') as file:
        reader = csv.reader(file)
        rows = 0
        for row in reader:
            rows += 1
    return rows


def merge_data(worksheet, stock_type, raw_data_folder, f_name):
    last_row = 0
    csv_row_count = 0
    row_init = 1
    csv_rows_tmp = []

    with open('{}/{}'.format(raw_data_folder, f_name), 'r', encoding='utf-8') as file:
        spamReader = csv.reader(file, delimiter=',', quotechar='"')
        # read data to memory and filter out invalid data
        count = 0
        for row in spamReader:
            count += 1
            if count > 2:  # column 1 and 2 are headers, ignore it
                if (stock_type == 'TSE'):
                    if (row[2] != '--' or row[6] != ''):
                        csv_rows_tmp.append(row)
                        last_row += 1
                elif (stock_type == 'OTC'):
                    if (row[1] != '0') and (row[2] != '---' or row[6] != '---'):
                        csv_rows_tmp.append(row)
                        last_row += 1

    if last_row < 248:
        last_row = 248

    for row in csv_rows_tmp:
        csv_row_count += 1
        if csv_row_count > (last_row - 248):
            try:
                new_open = float(row[2])
                new_high = float(row[3])
                new_low = float(row[4])
                new_close = float(row[5])
            except ValueError:
                new_open = row[2]
                new_high = row[3]
                new_low = row[4]
                new_close = row[5]

            worksheet.write(row_init, 0, row[0])
            worksheet.write(row_init, 1, int(row[1]) / 1000)
            worksheet.write(row_init, 2, new_open)
            worksheet.write(row_init, 3, new_high)
            worksheet.write(row_init, 4, new_low)
            worksheet.write(row_init, 5, new_close)
            row_init += 1


def summary_data_fill(summary_sheet, raw_data_folder, f_short_name, summary_idx):
    with open('{}/{}.csv'.format(raw_data_folder, f_short_name), 'r', encoding='utf-8') as file:
        reader = csv.reader(file)
        for row in reader:
            row[0] = ''.join(row[0].split())  # remove spaces
            stock_no = '[{:>6}]'.format(f_short_name)  # right-aligned with 6 spaces
            stock_name = u'{:<8} {}'.format(stock_no, row[0])
            break

    index = summary_idx + 2
    # stock_name headers
    summary_sheet.write(index, 0,
                        u"=HYPERLINK(\"[./{0}/{1}.xlsx]{1}!$A268\",\"{2}\")".format(FOLDER_SAVE,
                                                                                    f_short_name,
                                                                                    stock_name))
    # data value
    for idx in range(10):
        summary_sheet.write(index, 1 + idx, "='{}{}\[{}.xlsx]{}'!W{}".format(WorkingDirectory,
                                                                             FOLDER_SAVE, f_short_name,
                                                                             f_short_name,
                                                                             str(249 - idx)))  # start at B2


def process(stock_type, name, total, count):
    print("%s, create table [%6s]..." % (stock_type, name),
          "%3d" % (count * 100 / total) + "%\r")


def main():
    global WorkingDirectory
    WorkingDirectory = os.getcwd()

    # Make directory if not exist
    if not os.path.isdir(FOLDER_SAVE):
        os.mkdir(FOLDER_SAVE)

    # total files in both TSE/OTC
    tse_files = sum(
        os.path.isfile(os.path.join('{}/'.format(TSE_RAW_DATA_FOLDER), f)) for f in
        os.listdir('{}/'.format(TSE_RAW_DATA_FOLDER)))
    otc_files = sum(
        os.path.isfile(os.path.join('{}/'.format(OTC_RAW_DATA_FOLDER), f)) for f in
        os.listdir('{}/'.format(OTC_RAW_DATA_FOLDER)))

    # create summary worksheet
    tse_summary_sheet = summarybook.add_worksheet('TSE')
    tse_summary_sheet.freeze_panes(2, 1)
    tse_summary_sheet.merge_range('A1:L1', "", format)
    tse_summary_sheet.set_column(0, 0, 16, link_format)  # set column width
    tse_summary_sheet.set_row('B2,L2', None, format)
    tse_summary_sheet.write('A1', u'壓力值', format)
    tse_summary_sheet.write('A2', u'股票代號', format)
    tse_summary_sheet.set_column('B:Z', 13)  # set column width
    tse_summary_sheet.conditional_format('B3:L35800',
                                         {'type': 'cell', 'criteria': '<=', 'value': 0.6, 'format': summary_format})

    otc_summary_sheet = summarybook.add_worksheet('OTC')
    otc_summary_sheet.freeze_panes(2, 1)
    otc_summary_sheet.merge_range('A1:L1', "", format)
    otc_summary_sheet.set_column(0, 0, 16, link_format)  # set column width
    otc_summary_sheet.set_row('B2,L2', None, format)
    otc_summary_sheet.write('A1', u'壓力值', format)
    otc_summary_sheet.write('A2', u'股票代號', format)
    otc_summary_sheet.set_column('B:Z', 13)  # set column width
    otc_summary_sheet.conditional_format('B3:L35800',
                                         {'type': 'cell', 'criteria': '<=', 'value': 0.6, 'format': summary_format})
    # Add a button tied to a macro in the VBA project.
    tse_summary_sheet.insert_button('A1', {'macro': 'UpdateTableValues',
                                            'caption': 'Click to update!',
                                            'width': 200,
                                            'height': 20})

    stock_type_list = ['TSE', 'OTC']
    total_files_list = [tse_files, otc_files]
    summary_sheet_list = [tse_summary_sheet, otc_summary_sheet]
    summary_date_create = False
    summary_idx = 0

    for stock_type, raw_data, summary_sheet, total_files in zip(stock_type_list, RAW_DATA_FOLDER, summary_sheet_list,
                                                                 total_files_list):
        # create stocks worksheet
        for filename in sorted(glob.glob("./{}/*.csv".format(raw_data))):
            (f_path, f_name) = os.path.split(filename)
            (f_short_name, f_extension) = os.path.splitext(f_name)

            # print '%6s'%(f_short_name) + " creating table...",

            # create uniform worksheets
            spreadbook = xlsxwriter.Workbook('./{}/{}.xlsx'.format(FOLDER_SAVE, f_short_name))
            spread_format = spreadbook.add_format()
            spread_format.set_border(1)
            spread_format.set_bg_color('#cccccc')
            spread_format.set_align('center')
            spreadsheet = spreadbook.add_worksheet(f_short_name)
            spreadsheet.write_row('A1', titles, spread_format)
            spreadsheet.set_column(0, 0, 15)  # A1: set column width to 15
            spreadsheet.set_column(6, 6, 11)  # G1: set column width to 15
            spreadsheet.set_column(13, 14, 11)  # N1,O1: set column width to 15
            spreadsheet.set_column(16, 17, 15)  # Q1,R1: set column width to 15
            spreadsheet.set_column(19, 20, 13)  # T1,U1: set column width to 15
            spreadsheet.set_column(22, 22, 17)  # W1: set column width to 15

            # formula
            formula(spreadsheet)

            # chart
            chart_def(spreadbook, spreadsheet, f_short_name)

            # merge data
            merge_data(spreadsheet, stock_type, raw_data, f_name)

            # write and close
            spreadbook.close()

            # fill data at summary sheet
            if (summary_date_create == False):
                for idx in range(0, 10, 1):
                    tse_summary_sheet.write(1, 1 + idx, "='{}\{}\[{}.xlsx]{}'!A{}".format(WorkingDirectory,
                                                                                             FOLDER_SAVE,
                                                                                             f_short_name,
                                                                                             f_short_name,
                                                                                             str(249 - idx)),
                                             format)  # start at B2
                    otc_summary_sheet.write(1, 1 + idx, "='{}\{}\[{}.xlsx]{}'!A{}".format(WorkingDirectory,
                                                                                             FOLDER_SAVE,
                                                                                             f_short_name,
                                                                                             f_short_name,
                                                                                             str(249 - idx)),
                                             format)  # start at B2
                summary_date_create = True

            summary_data_fill(summary_sheet, raw_data, f_short_name, summary_idx)
            summary_idx += 1

    # Save summary book
    summarybook.close()
    print("Done!")


if __name__ == '__main__':
    main()

