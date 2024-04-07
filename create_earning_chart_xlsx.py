#coding=utf-8
import xlsxwriter
import glob
import csv
import os
import numpy as np

from os import mkdir
from os.path import isdir
from xlsxwriter.utility import xl_rowcol_to_cell
from datetime import datetime, timedelta

WorkingDirectory = os.getcwd()
TSE_RAW_DATA_FOLDER = 'tse_earning_raw_data'
OTC_RAW_DATA_FOLDER = 'otc_earning_raw_data'
OUTPUT_CHART_FILE = 'tse_otc_earning_chart.xlsx'
TOTAL_YEARS = 4 

Q4 = np.array([0,0,0], int) 

def formula(worksheet):

    global TOTAL_DAYS

    row_init = 0
    col_init = 0
    row_end = TOTAL_DAYS + 1 

    for row in range(5, row_end, 1):
        #current row value
        col_B = xl_rowcol_to_cell(row, 1)
        col_C = xl_rowcol_to_cell(row, 2)
        col_D = xl_rowcol_to_cell(row, 3)
        col_E = xl_rowcol_to_cell(row, 4)
        col_G = xl_rowcol_to_cell(row, 6)
        col_H = xl_rowcol_to_cell(row, 7)
        col_I = xl_rowcol_to_cell(row, 8)

        # 5MA
        first_row = xl_rowcol_to_cell(row-4, 4) #col_E
        worksheet.write_formula(row, col_init+6, "AVERAGE(%s:%s)"%(first_row, col_E))

        if row >= 30:
            # 30MA
            first_row = xl_rowcol_to_cell(row-29, 4) #col_E
            worksheet.write_formula(row, col_init+7, "AVERAGE(%s:%s)"%(first_row, col_E))

            # 5-30MA
            worksheet.write_formula(row, col_init+8, "%s-%s"%(col_G, col_H))
 
def chart_trend(market, spreadbook, worksheet, start_year, stock_count_row):

    # chart
    line_chart = spreadbook.add_chart({'type': 'line'})
    bar_chart = spreadbook.add_chart({'type': 'column'})
    line_yoy_chart = spreadbook.add_chart({'type': 'line'})
    bar_yoy_p_chart = spreadbook.add_chart({'type': 'column'})
    bar_yoy_op_chart = spreadbook.add_chart({'type': 'column'})
    bar_yoy_eps_chart = spreadbook.add_chart({'type': 'column'})
    # row to present
    revenue    = stock_count_row + 2 # 營收
    profit     = stock_count_row + 3 # 毛利
    profit_m   = stock_count_row + 4 # 毛利率
    o_profit   = stock_count_row + 5 # 營益
    o_profit_m = stock_count_row + 6 # 營益率
    eps        = stock_count_row + 7 # eps

    last_col_chr = chr(ord('C')+(TOTAL_YEARS*4-1))  # character for the last column

    ## data item
    # 營收
    line_chart.add_series({
        'name':       '=%s!$B2'%(market),
        'categories': '=%s!$C$1:$%c$1'%(market, last_col_chr),
        'values':     '=%s!$C$%s:$%c$%s'%(market,revenue, last_col_chr ,revenue),
        'marker':     {'type': 'diamond', 'size': 3},
    })
    # 毛利
    line_chart.add_series({
        'name':       '=%s!$B3'%(market),
        'categories': '=%s!$C$1:$%c$1'%(market, last_col_chr),
        'values':     '=%s!$C$%s:$%c$%s'%(market,profit, last_col_chr, profit),
    })
    # 營益 
    line_chart.add_series({
        'name':       '=%s!$B5'%(market),
        'categories': '=%s!$C$1:$%c$1'%(market, last_col_chr),
        'values':     '=%s!$C$%s:$%c$%s'%(market,o_profit, last_col_chr, o_profit),
    })

    # EPS
    bar_chart.add_series({
        'name':       '=%s!$B7'%(market),
        'categories': '=%s!$C$1:$%c$1'%(market, last_col_chr),
        'values':     '=%s!$C$%s:$%c$%s'%(market,eps , last_col_chr, eps),
        'y2_axis':    1,
    })

    line_yoy_chart.add_series({
        'name':       'YoY% - 營收',
        'categories': '=%s!$C$1:$%c$1'%(market, last_col_chr),
        'values':     '=%s!$C$%s:$%c$%s'%(market, revenue+3 , last_col_chr, revenue+3),
    })
    line_yoy_chart.add_series({
        'name':       'YoY% - 營益',
        'categories': '=%s!$C$1:$%c$1'%(market, last_col_chr),
        'values':     '=%s!$C$%s:$%c$%s'%(market, o_profit+3 , last_col_chr, o_profit+3),
    })
    line_yoy_chart.add_series({
        'name':       'YoY% - EPS',
        'categories': '=%s!$C$1:$%c$1'%(market, last_col_chr),
        'values':     '=%s!$C$%s:$%c$%s'%(market, eps+3 , last_col_chr, eps+3),
    })

    # EPS, YoY%
    bar_yoy_eps_chart.add_series({
        'name':       'YoY% - EPS',
        'categories': '=%s!$C$1:$%c$1'%(market, last_col_chr),
        'values':     '=%s!$C$%s:$%c$%s'%(market, eps+3 , last_col_chr, eps+3),
    })

    #毛利率
    bar_yoy_p_chart.add_series({
        'name':       '=%s!$B4'%(market),
        'categories': '=%s!$C$1:$%c$1'%(market, last_col_chr),
        'values':     '=%s!$C$%s:$%c$%s'%(market, profit_m, last_col_chr, profit_m),
    })
    #營益率
    bar_yoy_op_chart.add_series({
        'name':       '=%s!$B6'%(market),
        'categories': '=%s!$C$1:$%c$1'%(market, last_col_chr),
        'values':     '=%s!$C$%s:$%c$%s'%(market, o_profit_m, last_col_chr, o_profit_m),
    })

    worksheet.insert_chart('A%d'%(stock_count_row+8), line_chart)
    worksheet.insert_chart('A%d'%(stock_count_row+20), bar_chart)
    worksheet.insert_chart('A%d'%(stock_count_row+32), line_yoy_chart)
    worksheet.insert_chart('A%d'%(stock_count_row+44), bar_yoy_p_chart)
    worksheet.insert_chart('A%d'%(stock_count_row+56), bar_yoy_op_chart)
    worksheet.insert_chart('A%d'%(stock_count_row+68), bar_yoy_eps_chart)


def main():

    global TOTAL_YEARS
    global Q4

    print("=== Generating %s Chart ==="%(OUTPUT_CHART_FILE))
    print("Working directory: %s"%(WorkingDirectory))
    ## load config
    now = datetime.now()
    year = now.year
    month = now.month
    year_range = range(year-TOTAL_YEARS, year)

    ## prepare output
    spreadbook = xlsxwriter.Workbook(OUTPUT_CHART_FILE)

    ## loop thru TSE, OTC
    for market in ['TSE', 'OTC']:
        # 產業別
        sector = None
        tse_otc_raw_data_folder = TSE_RAW_DATA_FOLDER if market == 'TSE' else OTC_RAW_DATA_FOLDER
        files = glob.glob('%s/%s/%s*'%(WorkingDirectory, tse_otc_raw_data_folder, market))
        if len(files) == 0:
            print("No file found in %s"%(tse_otc_raw_data_folder))
            continue
        else:
            print("%d files found in %s"%(len(files), tse_otc_raw_data_folder))

        # add worksheet
        worksheet = spreadbook.add_worksheet(market)
        row_count = 0
        worksheet.write('A1', '年度')
        worksheet.write('B1', '產業別')
        worksheet.write('C1', '股票代號')
        worksheet.write('D1', '公司名稱')
        worksheet.write('E1', '日期')
        worksheet.write('F1', '營收')
        worksheet.write('G1', '毛利')
        worksheet.write('H1', '毛利率')
        worksheet.write('I1', '營益')
        worksheet.write('J1', '營益率')
        worksheet.write('K1', 'EPS')

        stock_count_row = 1

        for filename in files:
            #print("loading: %s"%(filename))
            with open(filename, newline='', encoding='utf-8') as csvfile:
                spamreader = csv.reader(csvfile, delimiter=',')
                for i, row in enumerate(spamreader):
                    if i == 0: # skip the first row
                        continue
                    # check if it is Q4
                    if i == 1: # 檢查第一筆, 代表是否為Q4
                        if row[0] == "Q4":
                            Q4[market == 'TSE'] += 1
                            print("Detect Q4 for %s, total %d"%(market, Q4[market == 'TSE']))
                    stock_id = row[0]
                    date = row[1].split("/")
                    date = "%s%s%s"%(date[0], date[1], date[2])
                    revenue = row[2]
                    profit = row[3]
                    profit_m = row[4]
                    o_profit = row[5]
                    o_profit_m = row[6]
                    eps = row[7]
                    if sector is None: # first item only
                        sector = row[8]
                    worksheet.write(stock_count_row, 0, year_range[len(files)-1])
                    worksheet.write(stock_count_row, 1, sector)
                    worksheet.write(stock_count_row, 2, stock_id)
                    worksheet.write(stock_count_row, 3, row[9])
                    worksheet.write(stock_count_row, 4, date)
                    worksheet.write(stock_count_row, 5, float(revenue))
                    worksheet.write(stock_count_row, 6, float(profit))
                    worksheet.write(stock_count_row, 7, float(profit_m.strip('%'))/100)
                    worksheet.write(stock_count_row, 8, float(o_profit))
                    worksheet.write(stock_count_row, 9, float(o_profit_m.strip('%'))/100)
                    worksheet.write(stock_count_row, 10, float(eps))
                    stock_count_row += 1

        TOTAL_DAYS = stock_count_row
        print("TOTAL_DAYS = %d"%(TOTAL_DAYS))
        formula(worksheet)
        chart_trend(market, spreadbook, worksheet, year_range[len(files)-1], stock_count_row)

    spreadbook.close()

if __name__ == "__main__":
    main()

