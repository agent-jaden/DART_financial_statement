#-*- coding:utf-8 -*-
# Parsing dividends data from DART
import urllib.request
import urllib.parse
import xlsxwriter
import os
import time
import sys
import getopt
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
import re
import xlrd
import fix_yahoo_finance as yf
#from pandas_datareader import data
import pandas_datareader
import numpy as np
import matplotlib.pyplot as plt

#
def find_value(text, unit):
	return int(text.replace(" ","").replace("△","-").replace("(-)","-").replace("(","-").replace(")","").replace(",","").replace("=",""))/unit

# Draw figure of cashflows.
def draw_figure(income_list, income_list2, year_list, op_cashflow_list, fcf_list, div_list, stock_close):
	
	for i in range(len(income_list)):
		if income_list[i] == 0.0:
			income_list[i] = income_list2[i]

	fig, ax1 = plt.subplots()

	ax1.plot(year_list, op_cashflow_list, label="Op Cashflow", color='r', marker='D')
	ax1.plot(year_list, fcf_list, label="Free Cashflow", color='y', marker='D')
	ax1.plot(year_list, income_list, label="Net Income", color='b', marker='D')
	ax1.plot(year_list, div_list, label="Dividends", color='g', marker='D')
	#ax1.plot(year_list, cash_equivalents_list, label="Cash & Cash Equivalents", color='magenta', marker='D', linestyle ='dashed')
	ax1.set_xlabel("YEAR")
	plt.legend(loc=2)

	ax2 = ax1.twinx().twiny()
	ax2.plot(stock_close, label="Stock Price", color='gray')

	#plt.title(corp)
	plt.legend(loc=4)
	plt.show()

# Write financial statements to Excel file.
def write_excel_file(workbook_name, dart_div_list, cashflow_list, balance_sheet_list, income_statement_list, corp, stock_code, stock_cat):
	# Write an Excel file

	#workbook = xlsxwriter.Workbook(workbook_name)
	#if os.path.isfile(os.path.join(cur_dir, workbook_name)):
	#	os.remove(os.path.join(cur_dir, workbook_name))
	workbook = xlsxwriter.Workbook(workbook_name)

	worksheet_result = workbook.add_worksheet('DART사업보고서')
	filter_format = workbook.add_format({'bold':True,
										'fg_color': '#D7E4BC'
										})
	filter_format2 = workbook.add_format({'bold':True
										})

	percent_format = workbook.add_format({'num_format': '0.00%'})

	roe_format = workbook.add_format({'bold':True,
									  'underline': True,
									  'num_format': '0.00%'})

	num_format = workbook.add_format({'num_format':'0.00'})
	num2_format = workbook.add_format({'num_format':'#,##0'})
	num3_format = workbook.add_format({'num_format':'#,##0.00',
									  'fg_color':'#FCE4D6'})

	worksheet_result.set_column('A:A', 10)
	worksheet_result.set_column('B:B', 15)
	worksheet_result.set_column('C:C', 15)
	worksheet_result.set_column('D:D', 20)
	worksheet_result.set_column('H:H', 15)
	worksheet_result.set_column('I:I', 15)
	worksheet_result.set_column('J:J', 15)
	worksheet_result.set_column('K:K', 15)

	worksheet_result.write(0, 0, "날짜", filter_format)
	worksheet_result.write(0, 1, "회사명", filter_format)
	worksheet_result.write(0, 2, "분류", filter_format)
	worksheet_result.write(0, 3, "제목", filter_format)
	worksheet_result.write(0, 4, "link", filter_format)
	worksheet_result.write(0, 5, "결산년도", filter_format)
	worksheet_result.write(0, 6, "영업활동 현금흐름", filter_format)
	worksheet_result.write(0, 7, "영업에서 창출된 현금흐름", filter_format)
	worksheet_result.write(0, 8, "당기순이익", filter_format)
	worksheet_result.write(0, 9, "투자활동 현금흐름", filter_format)
	worksheet_result.write(0, 10, "유형자산의 취득", filter_format)
	worksheet_result.write(0, 11, "무형자산의 취득", filter_format)
	worksheet_result.write(0, 12, "토지의 취득", filter_format)
	worksheet_result.write(0, 13, "건물의 취득", filter_format)
	worksheet_result.write(0, 14, "구축물의 취득", filter_format)
	worksheet_result.write(0, 15, "기계장치의 취득", filter_format)
	worksheet_result.write(0, 16, "건설중인자산의 증가", filter_format)
	worksheet_result.write(0, 17, "차량운반구의 취득", filter_format)
	worksheet_result.write(0, 18, "비품의 취득", filter_format)
	worksheet_result.write(0, 19, "공구기구의 취득", filter_format)
	worksheet_result.write(0, 20, "시험 연구 설비의 취득", filter_format)
	worksheet_result.write(0, 21, "렌탈 자산의 취득", filter_format)
	worksheet_result.write(0, 22, "영업권의 취득", filter_format)
	worksheet_result.write(0, 23, "산업재산권의 취득", filter_format)
	worksheet_result.write(0, 24, "소프트웨어의 취득", filter_format)
	worksheet_result.write(0, 25, "기타의무형자산의 취득", filter_format)
	worksheet_result.write(0, 26, "투자부동산의 취득", filter_format)
	worksheet_result.write(0, 27, "관계기업투자의 취득", filter_format)
	worksheet_result.write(0, 28, "재무활동 현금흐름", filter_format)
	worksheet_result.write(0, 29, "단기차입금의 증가", filter_format)
	worksheet_result.write(0, 30, "배당금 지급", filter_format)
	worksheet_result.write(0, 31, "자기주식의 취득", filter_format)
	worksheet_result.write(0, 32, "기초현금 및 현금성자산", filter_format)
	worksheet_result.write(0, 33, "기말현금 및 현금성자산", filter_format)

	for k in range(len(dart_div_list)):
		worksheet_result.write(k+1,0, dart_div_list[k][0], num2_format)
		worksheet_result.write(k+1,1, dart_div_list[k][1], num2_format)
		worksheet_result.write(k+1,2, dart_div_list[k][2], num2_format)
		worksheet_result.write(k+1,3, dart_div_list[k][3], num2_format)
		worksheet_result.write(k+1,4, dart_div_list[k][4], num2_format)
		worksheet_result.write(k+1,5, cashflow_list[k]	['year']					, num2_format)
		worksheet_result.write(k+1,6, cashflow_list[k]	['op_cashflow']				, num2_format)
		worksheet_result.write(k+1,7, cashflow_list[k]	['op_cashflow_sub1']		, num2_format)
		worksheet_result.write(k+1,8, cashflow_list[k]	['op_cashflow_sub2']		, num2_format)
		worksheet_result.write(k+1,9, cashflow_list[k]	['invest_cashflow']			, num2_format)
		worksheet_result.write(k+1,10, cashflow_list[k]	['invest_cashflow_sub1']	, num2_format)
		worksheet_result.write(k+1,11, cashflow_list[k]	['invest_cashflow_sub2']	, num2_format)
		worksheet_result.write(k+1,12, cashflow_list[k]	['invest_cashflow_sub3']	, num2_format)
		worksheet_result.write(k+1,13, cashflow_list[k]	['invest_cashflow_sub4']	, num2_format)
		worksheet_result.write(k+1,14, cashflow_list[k]	['invest_cashflow_sub5']	, num2_format)
		worksheet_result.write(k+1,15, cashflow_list[k]	['invest_cashflow_sub6']	, num2_format)
		worksheet_result.write(k+1,16, cashflow_list[k]	['invest_cashflow_sub7']	, num2_format)
		worksheet_result.write(k+1,17, cashflow_list[k]	['invest_cashflow_sub8']	, num2_format)
		worksheet_result.write(k+1,18, cashflow_list[k]	['invest_cashflow_sub9']	, num2_format)
		worksheet_result.write(k+1,19, cashflow_list[k]	['invest_cashflow_sub10']	, num2_format)
		worksheet_result.write(k+1,20, cashflow_list[k]	['invest_cashflow_sub11']	, num2_format)
		worksheet_result.write(k+1,21, cashflow_list[k]	['invest_cashflow_sub12']	, num2_format)
		worksheet_result.write(k+1,22, cashflow_list[k]	['invest_cashflow_sub13']	, num2_format)
		worksheet_result.write(k+1,23, cashflow_list[k]	['invest_cashflow_sub14']	, num2_format)
		worksheet_result.write(k+1,24, cashflow_list[k]	['invest_cashflow_sub15']	, num2_format)
		worksheet_result.write(k+1,25, cashflow_list[k]	['invest_cashflow_sub16']	, num2_format)
		worksheet_result.write(k+1,26, cashflow_list[k]	['invest_cashflow_sub17']	, num2_format)
		worksheet_result.write(k+1,27, cashflow_list[k]	['invest_cashflow_sub18']	, num2_format)
		worksheet_result.write(k+1,28, cashflow_list[k]	['fin_cashflow']			, num2_format)
		worksheet_result.write(k+1,29, cashflow_list[k]	['fin_cashflow_sub1']		, num2_format)
		worksheet_result.write(k+1,30, cashflow_list[k]	['fin_cashflow_sub2']		, num2_format)
		worksheet_result.write(k+1,31, cashflow_list[k]	['fin_cashflow_sub3']		, num2_format)
		worksheet_result.write(k+1,32, cashflow_list[k]	['start_cash']				, num2_format)
		worksheet_result.write(k+1,33, cashflow_list[k]	['end_cash']				, num2_format)

	cashflow_list.reverse() 
	worksheet_cashflow = workbook.add_worksheet('cashflow')
	
	worksheet_cashflow.set_column('A:A', 30)
	worksheet_cashflow.write(0, 0, "결산년도", filter_format)
	worksheet_cashflow.write(1, 0, "영업활동 현금흐름", filter_format)
	worksheet_cashflow.write(2, 0, "영업에서 창출된 현금흐름", filter_format2)
	worksheet_cashflow.write(3, 0, "당기순이익", filter_format2)
	worksheet_cashflow.write(4, 0, "감가상각비", filter_format2)
	worksheet_cashflow.write(5, 0, "신탁계정대", filter_format2)
	worksheet_cashflow.write(6, 0, "투자활동 현금흐름", filter_format)
	worksheet_cashflow.write(7, 0, "유형자산의 취득", filter_format2)
	worksheet_cashflow.write(8, 0, "무형자산의 취득", filter_format2)
	worksheet_cashflow.write(9, 0, "토지의 취득", filter_format2)
	worksheet_cashflow.write(10, 0, "건물의 취득", filter_format2)
	worksheet_cashflow.write(11, 0, "구축물의 취득", filter_format2)
	worksheet_cashflow.write(12, 0, "기계장치의 취득", filter_format2)
	worksheet_cashflow.write(13, 0, "건설중인자산의 증가", filter_format2)
	worksheet_cashflow.write(14, 0, "차량운반구의 취득", filter_format2)
	worksheet_cashflow.write(15, 0, "비품의 취득", filter_format2)
	worksheet_cashflow.write(16, 0, "공구기구의 취득", filter_format2)
	worksheet_cashflow.write(17, 0, "시험 연구 설비의 취득", filter_format2)
	worksheet_cashflow.write(18, 0, "렌탈 자산의 취득", filter_format2)
	worksheet_cashflow.write(19, 0, "영업권의 취득", filter_format2)
	worksheet_cashflow.write(20, 0, "산업재산권의 취득", filter_format2)
	worksheet_cashflow.write(21, 0, "소프트웨어의 취득", filter_format2)
	worksheet_cashflow.write(22, 0, "기타의무형자산의 취득", filter_format2)
	worksheet_cashflow.write(23, 0, "투자부동산의 취득", filter_format2)
	worksheet_cashflow.write(24, 0, "관계기업투자의 취득", filter_format2)
	worksheet_cashflow.write(25, 0, "재무활동 현금흐름", filter_format)
	worksheet_cashflow.write(26, 0, "단기차입금의 증가", filter_format2)
	worksheet_cashflow.write(27, 0, "배당금 지급", filter_format2)
	worksheet_cashflow.write(28, 0, "자기주식의 취득", filter_format2)
	worksheet_cashflow.write(29, 0, "기초현금 및 현금성자산", filter_format)
	worksheet_cashflow.write(30, 0, "기말현금 및 현금성자산", filter_format)
	worksheet_cashflow.write(31, 0, "당기순이익 손익계산서", filter_format2)
	worksheet_cashflow.write(32, 0, "잉여현금흐름(FCF)", filter_format)

	prev_year = 0
	j = 0

	year_list = []
	op_cashflow_list = []
	fcf_list = []
	income_list = []
	income_list2 = []
	div_list = []
	cash_equivalents_list = []

	for k in range(len(cashflow_list)):
		fcf = cashflow_list[k]['op_cashflow']
		fcf = fcf - abs(cashflow_list[k]['invest_cashflow_sub1'])
		fcf = fcf - abs(cashflow_list[k]['invest_cashflow_sub2'])
		fcf = fcf - abs(cashflow_list[k]['invest_cashflow_sub3'])
		fcf = fcf - abs(cashflow_list[k]['invest_cashflow_sub4'])
		fcf = fcf - abs(cashflow_list[k]['invest_cashflow_sub5'])
		fcf = fcf - abs(cashflow_list[k]['invest_cashflow_sub6'])
		fcf = fcf - abs(cashflow_list[k]['invest_cashflow_sub7'])
		fcf = fcf - abs(cashflow_list[k]['invest_cashflow_sub8'])
		fcf = fcf - abs(cashflow_list[k]['invest_cashflow_sub9'])
		fcf = fcf - abs(cashflow_list[k]['invest_cashflow_sub10'])
		fcf = fcf - abs(cashflow_list[k]['invest_cashflow_sub11'])
		fcf = fcf - abs(cashflow_list[k]['invest_cashflow_sub12'])
		fcf = fcf - abs(cashflow_list[k]['invest_cashflow_sub13'])
		fcf = fcf - abs(cashflow_list[k]['invest_cashflow_sub14'])
		fcf = fcf - abs(cashflow_list[k]['invest_cashflow_sub15'])
		fcf = fcf - abs(cashflow_list[k]['invest_cashflow_sub16'])
	
		if cashflow_list[k]['op_cashflow_sub1'] != "FINDING LINE NUMBER ERROR":
			# Overwirting
			if prev_year == cashflow_list[k]['year']:
				worksheet_cashflow.write(0, j, str(cashflow_list[k]['year'])+"년")
				worksheet_cashflow.write(1, j, cashflow_list[k]['op_cashflow']				, num2_format)
				worksheet_cashflow.write(2, j, cashflow_list[k]['op_cashflow_sub1']			, num2_format)
				worksheet_cashflow.write(3, j, cashflow_list[k]['op_cashflow_sub2']			, num2_format)
				worksheet_cashflow.write(4, j, cashflow_list[k]['op_cashflow_sub3']			, num2_format)
				worksheet_cashflow.write(5, j, cashflow_list[k]['op_cashflow_sub4']			, num2_format)
				worksheet_cashflow.write(6, j, cashflow_list[k]['invest_cashflow']			, num2_format)
				worksheet_cashflow.write(7, j, cashflow_list[k]['invest_cashflow_sub1']		, num2_format)
				worksheet_cashflow.write(8, j, cashflow_list[k]['invest_cashflow_sub2']		, num2_format)
				worksheet_cashflow.write(9, j, cashflow_list[k]['invest_cashflow_sub3']		, num2_format)
				worksheet_cashflow.write(10, j, cashflow_list[k]['invest_cashflow_sub4']		, num2_format)
				worksheet_cashflow.write(11, j, cashflow_list[k]['invest_cashflow_sub5']		, num2_format)
				worksheet_cashflow.write(12, j, cashflow_list[k]['invest_cashflow_sub6']	, num2_format)
				worksheet_cashflow.write(13, j, cashflow_list[k]['invest_cashflow_sub7']	, num2_format)
				worksheet_cashflow.write(14, j, cashflow_list[k]['invest_cashflow_sub8']	, num2_format)
				worksheet_cashflow.write(15, j, cashflow_list[k]['invest_cashflow_sub9']	, num2_format)
				worksheet_cashflow.write(16, j, cashflow_list[k]['invest_cashflow_sub10']	, num2_format)
				worksheet_cashflow.write(17, j, cashflow_list[k]['invest_cashflow_sub11']	, num2_format)
				worksheet_cashflow.write(18, j, cashflow_list[k]['invest_cashflow_sub12']	, num2_format)
				worksheet_cashflow.write(19, j, cashflow_list[k]['invest_cashflow_sub13']	, num2_format)
				worksheet_cashflow.write(20, j, cashflow_list[k]['invest_cashflow_sub14']	, num2_format)
				worksheet_cashflow.write(21, j, cashflow_list[k]['invest_cashflow_sub15']	, num2_format)
				worksheet_cashflow.write(22, j, cashflow_list[k]['invest_cashflow_sub16']	, num2_format)
				worksheet_cashflow.write(23, j, cashflow_list[k]['invest_cashflow_sub17']	, num2_format)
				worksheet_cashflow.write(24, j, cashflow_list[k]['invest_cashflow_sub18']	, num2_format)
				worksheet_cashflow.write(25, j, cashflow_list[k]['fin_cashflow']			, num2_format)
				worksheet_cashflow.write(26, j, cashflow_list[k]['fin_cashflow_sub1']		, num2_format)
				worksheet_cashflow.write(27, j, cashflow_list[k]['fin_cashflow_sub2']		, num2_format)
				worksheet_cashflow.write(28, j, cashflow_list[k]['fin_cashflow_sub3']		, num2_format)
				worksheet_cashflow.write(29, j, cashflow_list[k]['start_cash']				, num2_format)
				worksheet_cashflow.write(30, j, cashflow_list[k]['end_cash']				, num2_format)
				worksheet_cashflow.write(31, j, cashflow_list[k]['net_income']				, num2_format)
				worksheet_cashflow.write(32, j, fcf, num2_format)
			else:
				worksheet_cashflow.write(0, j+1, str(cashflow_list[k]['year'])+"년")
				worksheet_cashflow.write(1, j+1, cashflow_list[k]['op_cashflow']			, num2_format)
				worksheet_cashflow.write(2, j+1, cashflow_list[k]['op_cashflow_sub1']		, num2_format)
				worksheet_cashflow.write(3, j+1, cashflow_list[k]['op_cashflow_sub2']		, num2_format)
				worksheet_cashflow.write(4, j+1, cashflow_list[k]['op_cashflow_sub3']		, num2_format)
				worksheet_cashflow.write(5, j+1, cashflow_list[k]['op_cashflow_sub4']		, num2_format)
				worksheet_cashflow.write(6, j+1, cashflow_list[k]['invest_cashflow']		, num2_format)
				worksheet_cashflow.write(7, j+1, cashflow_list[k]['invest_cashflow_sub1']	, num2_format)
				worksheet_cashflow.write(8, j+1, cashflow_list[k]['invest_cashflow_sub2']	, num2_format)
				worksheet_cashflow.write(9, j+1, cashflow_list[k]['invest_cashflow_sub3']	, num2_format)
				worksheet_cashflow.write(10, j+1, cashflow_list[k]['invest_cashflow_sub4']	, num2_format)
				worksheet_cashflow.write(11, j+1, cashflow_list[k]['invest_cashflow_sub5']	, num2_format)
				worksheet_cashflow.write(12, j+1, cashflow_list[k]['invest_cashflow_sub6']	, num2_format)
				worksheet_cashflow.write(13, j+1, cashflow_list[k]['invest_cashflow_sub7']	, num2_format)
				worksheet_cashflow.write(14, j+1, cashflow_list[k]['invest_cashflow_sub8']	, num2_format)
				worksheet_cashflow.write(15, j+1, cashflow_list[k]['invest_cashflow_sub9']	, num2_format)
				worksheet_cashflow.write(16, j+1, cashflow_list[k]['invest_cashflow_sub10']	, num2_format)
				worksheet_cashflow.write(17, j+1, cashflow_list[k]['invest_cashflow_sub11']	, num2_format)
				worksheet_cashflow.write(18, j+1, cashflow_list[k]['invest_cashflow_sub12']	, num2_format)
				worksheet_cashflow.write(19, j+1, cashflow_list[k]['invest_cashflow_sub13']	, num2_format)
				worksheet_cashflow.write(20, j+1, cashflow_list[k]['invest_cashflow_sub14']	, num2_format)
				worksheet_cashflow.write(21, j+1, cashflow_list[k]['invest_cashflow_sub15']	, num2_format)
				worksheet_cashflow.write(22, j+1, cashflow_list[k]['invest_cashflow_sub16']	, num2_format)
				worksheet_cashflow.write(23, j+1, cashflow_list[k]['invest_cashflow_sub17']	, num2_format)
				worksheet_cashflow.write(24, j+1, cashflow_list[k]['invest_cashflow_sub18']	, num2_format)
				worksheet_cashflow.write(25, j+1, cashflow_list[k]['fin_cashflow']			, num2_format)
				worksheet_cashflow.write(26, j+1, cashflow_list[k]['fin_cashflow_sub1']		, num2_format)
				worksheet_cashflow.write(27, j+1, cashflow_list[k]['fin_cashflow_sub2']		, num2_format)
				worksheet_cashflow.write(28, j+1, cashflow_list[k]['fin_cashflow_sub3']		, num2_format)
				worksheet_cashflow.write(29, j+1, cashflow_list[k]['start_cash']			, num2_format)
				worksheet_cashflow.write(30, j+1, cashflow_list[k]['end_cash']				, num2_format)
				worksheet_cashflow.write(31, j+1, cashflow_list[k]['net_income']			, num2_format)
				worksheet_cashflow.write(32, j+1, fcf, num2_format)
			
				year_list.append(cashflow_list[k]['year'])
				op_cashflow_list.append(cashflow_list[k]['op_cashflow'])
				fcf_list.append(fcf)
				income_list.append(cashflow_list[k]['op_cashflow_sub2'])
				income_list2.append(cashflow_list[k]['net_income'])
				div_list.append(abs(cashflow_list[k]['fin_cashflow_sub2']))
				cash_equivalents_list.append(cashflow_list[k]['end_cash'])
				j = j+1
		
		prev_year = cashflow_list[k]['year']

	# Balance sheet
	balance_sheet_list.reverse() 
	worksheet_bs= workbook.add_worksheet('Balance Sheet')
	
	prev_year = 0
	j = 0
	
	worksheet_bs.set_column('A:A', 30)
	worksheet_bs.write(0, 0, "결산년도", filter_format)
	worksheet_bs.write(1, 0, "유동자산", filter_format)
	worksheet_bs.write(2, 0, "현금 및 현금성 자산", filter_format2)
	worksheet_bs.write(3, 0, "매출채권", filter_format2)
	worksheet_bs.write(4, 0, "재고자산", filter_format2)
	worksheet_bs.write(5, 0, "비유동자산", filter_format)
	worksheet_bs.write(6, 0, "유형자산", filter_format2)
	worksheet_bs.write(7, 0, "무형자산", filter_format2)
	worksheet_bs.write(8, 0, "자산총계", filter_format)
	worksheet_bs.write(9, 0, "유동부채", filter_format)
	worksheet_bs.write(10, 0, "매입채무", filter_format2)
	worksheet_bs.write(11, 0, "단기차입금", filter_format2)
	worksheet_bs.write(12, 0, "미지급금", filter_format2)
	worksheet_bs.write(13, 0, "비유동부채", filter_format)
	worksheet_bs.write(14, 0, "사채", filter_format2)
	worksheet_bs.write(15, 0, "장기차입금", filter_format2)
	worksheet_bs.write(16, 0, "장기미지급금", filter_format2)
	worksheet_bs.write(17, 0, "이연법인세부채", filter_format2)
	worksheet_bs.write(18, 0, "부채총계", filter_format)
	worksheet_bs.write(19, 0, "자본금", filter_format2)
	worksheet_bs.write(20, 0, "주식발행초과금", filter_format2)
	worksheet_bs.write(21, 0, "이익잉여금", filter_format2)
	worksheet_bs.write(22, 0, "자본총계", filter_format)
	
	for k in range(len(balance_sheet_list)):
		# Overwirting
		if prev_year == balance_sheet_list[k]['year']:
			w = j
		else:
			w = j+1

		worksheet_bs.write(0, w, str(balance_sheet_list[k]['year'])+"년")
		worksheet_bs.write(1, w, balance_sheet_list[k]['asset_current']				, num2_format)
		worksheet_bs.write(2, w, balance_sheet_list[k]['asset_current_sub1']			, num2_format)
		worksheet_bs.write(3, w, balance_sheet_list[k]['asset_current_sub2']			, num2_format)
		worksheet_bs.write(4, w, balance_sheet_list[k]['asset_current_sub3']			, num2_format)
		worksheet_bs.write(5, w, balance_sheet_list[k]['asset_non_current']			, num2_format)
		worksheet_bs.write(6, w, balance_sheet_list[k]['asset_non_current_sub1']		, num2_format)
		worksheet_bs.write(7, w, balance_sheet_list[k]['asset_non_current_sub2']		, num2_format)
		worksheet_bs.write(8, w, balance_sheet_list[k]['asset_sum']					, num2_format)
		worksheet_bs.write(9, w, balance_sheet_list[k]['liability_current']			, num2_format)
		worksheet_bs.write(10, w, balance_sheet_list[k]['liability_current_sub1']		, num2_format)
		worksheet_bs.write(11, w, balance_sheet_list[k]['liability_current_sub2']		, num2_format)
		worksheet_bs.write(12, w, balance_sheet_list[k]['liability_current_sub3']		, num2_format)
		worksheet_bs.write(13, w, balance_sheet_list[k]['liability_non_current']		, num2_format)
		worksheet_bs.write(14, w, balance_sheet_list[k]['liability_non_current_sub1']	, num2_format)
		worksheet_bs.write(15, w, balance_sheet_list[k]['liability_non_current_sub2']	, num2_format)
		worksheet_bs.write(16, w, balance_sheet_list[k]['liability_non_current_sub3']	, num2_format)
		worksheet_bs.write(17, w, balance_sheet_list[k]['liability_non_current_sub4']	, num2_format)
		worksheet_bs.write(18, w, balance_sheet_list[k]['liability_sum']				, num2_format)
		worksheet_bs.write(19, w, balance_sheet_list[k]['equity']						, num2_format)
		worksheet_bs.write(20, w, balance_sheet_list[k]['equity_sub1']				, num2_format)
		worksheet_bs.write(21, w, balance_sheet_list[k]['equity_sub2']				, num2_format)
		worksheet_bs.write(22, w, balance_sheet_list[k]['equity_sum']					, num2_format)
		
		j = j+1
		prev_year = balance_sheet_list[k]['year']

	# Income statement
	income_statement_list.reverse() 
	worksheet_income= workbook.add_worksheet('Income Statement')

	prev_year = 0
	j = 0
	
	worksheet_income.set_column('A:A', 30)
	worksheet_income.write(0, 0, "결산년도", filter_format)
	worksheet_income.write(1, 0, "매출액", filter_format)
	worksheet_income.write(2, 0, "매출원가", filter_format2)
	worksheet_income.write(3, 0, "매출총이익", filter_format2)
	worksheet_income.write(4, 0, "판매비와관리비", filter_format2)
	worksheet_income.write(5, 0, "영업이익", filter_format)
	worksheet_income.write(6, 0, "기타수익", filter_format2)
	worksheet_income.write(7, 0, "기타비용", filter_format2)
	worksheet_income.write(8, 0, "금융수익", filter_format2)
	worksheet_income.write(9, 0, "금융비용", filter_format2)
	worksheet_income.write(10, 0, "영업외수익", filter_format2)
	worksheet_income.write(11, 0, "영업외비용", filter_format2)
	worksheet_income.write(12, 0, "법인세비용차감전순이익", filter_format)
	worksheet_income.write(13, 0, "법인세비용", filter_format2)
	worksheet_income.write(14, 0, "당기순이익", filter_format)
	worksheet_income.write(15, 0, "기본주당이익", filter_format)

	for k in range(len(income_statement_list)):
		# Overwirting
		if prev_year == income_statement_list[k]['year']:
			w = j
		else:
			w = j+1

		worksheet_income.write(0, w, str(income_statement_list[k]['year'])+"년")
		worksheet_income.write(1, w, income_statement_list[k] ['sales']			, num2_format)
		worksheet_income.write(2, w, income_statement_list[k] ['sales_sub1']		, num2_format)
		worksheet_income.write(3, w, income_statement_list[k] ['sales_sub2']		, num2_format)
		worksheet_income.write(4, w, income_statement_list[k] ['sales_sub3']		, num2_format)
		worksheet_income.write(5, w, income_statement_list[k] ['op_income']		, num2_format)
		worksheet_income.write(6, w, income_statement_list[k] ['op_income_sub1']	, num2_format)
		worksheet_income.write(7, w, income_statement_list[k] ['op_income_sub2']	, num2_format)
		worksheet_income.write(8, w, income_statement_list[k] ['op_income_sub3']	, num2_format)
		worksheet_income.write(9, w, income_statement_list[k] ['op_income_sub4']	, num2_format)
		worksheet_income.write(10, w, income_statement_list[k]['op_income_sub6']	, num2_format)
		worksheet_income.write(11, w, income_statement_list[k]['op_income_sub7']	, num2_format)
		worksheet_income.write(12, w, income_statement_list[k]['op_income_sub5']	, num2_format)
		worksheet_income.write(13, w, income_statement_list[k]['tax']				, num2_format)
		worksheet_income.write(14, w, income_statement_list[k]['net_income']		, num2_format)
		worksheet_income.write(15, w, income_statement_list[k]['eps']				, num2_format)
		
		j = j+1
		prev_year = income_statement_list[k]['year']
	
	
	j = 0
	
	# Chart WORKSHEET	
	chart = workbook.add_chart({'type':'line'})
	chart.add_series({
					'categories':'=cashflow!$B$1:$Q$1',
					'name':'=cashflow!A2',
					'values':'=cashflow!$B$2:$Q$2',
					'marker':{'type': 'diamond'}
					})
	chart.add_series({
					'name':'=cashflow!A4',
					'values':'=cashflow!$B$4:$Q$4',
					'marker':{'type': 'diamond'}
					})
	chart.add_series({
					'name':'=cashflow!A26',
					'values':'=cashflow!$B$26:$Q$26',
					'marker':{'type': 'diamond'}
					})
	chart.set_legend({'font':{'bold':1}})
	chart.set_x_axis({'name':"결산년도"})
	chart.set_y_axis({'name':"단위:억원"})
	chart.set_title({'name':corp})

	worksheet_cashflow.insert_chart('C30', chart)

	old_year = cashflow_list[0]['year']

	if (stock_code != ""):
		yf.pdr_override()
		start_date = str(old_year)+'-01-01'
		if stock_cat == "코스피":
			ticker = stock_code+'.KS'
		else:
			ticker = stock_code+'.KQ'

		print("ticker", ticker)
		print("start date", start_date)
		stock_read = pandas_datareader.data.get_data_yahoo(ticker, start_date)
		stock_close = stock_read['Close'].values
		stock_datetime64 = stock_read.index.values

		stock_date = []

		for date in stock_datetime64:
			unix_epoch = np.datetime64(0, 's')
			one_second = np.timedelta64(1, 's')
			seconds_since_epoch = (date - unix_epoch) / one_second
			
			day = datetime.utcfromtimestamp(seconds_since_epoch)
			stock_date.append(day.strftime('%Y-%m-%d'))

		worksheet_stock = workbook.add_worksheet('stock_chart')

		worksheet_stock.write(0, 0, "date")
		worksheet_stock.write(0, 1, "Close")
		
		for i in range(len(stock_close)):
			worksheet_stock.write(i+1, 0, stock_date[i])
			worksheet_stock.write(i+1, 1, stock_close[i])
		
		chart = workbook.add_chart({'type':'line'})
		chart.add_series({
						'categories':'=stock_chart!$A$2:$A$'+str(len(stock_close)+1),
						'name':'=stock_chart!B1',
						'values':'=stock_chart!$B$2:$B$'+str(len(stock_close)+1)
						})

		worksheet_stock.insert_chart('D3', chart)

	workbook.close()
	draw_figure(income_list, income_list2, year_list, op_cashflow_list, fcf_list, div_list, stock_close)

# Get information of balance sheet
def scrape_balance_sheet(balance_sheet_table, year, unit):

	#유동자산
	##현금및현금성자산
	##매출채권
	##재고자산
	#비유동자산
	##유형자산
	##무형자산
	#자산총계
	#유동부채
	##매입채무
	##단기차입금
	##미지급금
	#비유동부채
	##사채
	##장기차입금
	##장기미지급금
	##이연법인세부채
	#부채총계
	##자본금
	##주식발행초과금
	##이익잉여금
	#자본총계

	re_asset_list = []

	re_asset_current				=	re.compile("^유[ \s]*동[ \s]*자[ \s]*산|\.[ \s]*유[ \s]*동[ \s]*자[ \s]*산")
	re_asset_current_sub1			=	re.compile("현[ \s]*금[ \s]*및[ \s]*현[ \s]*금[ \s]*((성[ \s]*자[ \s]*산)|(등[ \s]*가[ \s]*물))")
	re_asset_current_sub2			=	re.compile("매[ \s]*출[ \s]*채[ \s]*권")
	re_asset_current_sub3			=	re.compile("재[ \s]*고[ \s]*자[ \s]*산")
	re_asset_non_current			=	re.compile("비[ \s]*유[ \s]*동[ \s]*자[ \s]*산")
	re_asset_non_current_sub1		=	re.compile("유[ \s]*형[ \s]*자[ \s]*산")
	re_asset_non_current_sub2		=	re.compile("무[ \s]*형[ \s]*자[ \s]*산")
	re_asset_sum					=	re.compile("자[ \s]*산[ \s]*총[ \s]*계")
	re_liability_current			=	re.compile("^유[ \s]*동[ \s]*부[ \s]*채|\.[ \s]*유[ \s]*동[ \s]*부[ \s]*채")
	re_liability_current_sub1		=	re.compile("매[ \s]*입[ \s]*채[ \s]*무[ \s]*")
	re_liability_current_sub2		=	re.compile("단[ \s]*기[ \s]*차[ \s]*입[ \s]*금")
	re_liability_current_sub3		=	re.compile("^미[ \s]*지[ \s]*급[ \s]*금[ \s]*")
	re_liability_non_current		=	re.compile("비[ \s]*유[ \s]*동[ \s]*부[ \s]*채")
	re_liability_non_current_sub1	=	re.compile("사[ \s]*채[ \s]*")
	re_liability_non_current_sub2	=	re.compile("장[ \s]*기[ \s]*차[ \s]*입[ \s]*금")
	re_liability_non_current_sub3	=	re.compile("장[ \s]*기[ \s]*미[ \s]*지[ \s]*급[ \s]*금")
	re_liability_non_current_sub4	=	re.compile("이[ \s]*연[ \s]*법[ \s]*인[ \s]*세[ \s]*부[ \s]*채")
	re_liability_sum				=	re.compile("부[ \s]*채[ \s]*총[ \s]*계")
	re_equity						=	re.compile("자[ \s]*본[ \s]*금")
	re_equity_sub1					=	re.compile("주[ \s]*식[ \s]*발[ \s]*행[ \s]*초[ \s]*과[ \s]*금")
	re_equity_sub2					=	re.compile("이[ \s]*익[ \s]*잉[ \s]*여[ \s]*금")
	re_equity_sum					=	re.compile("자[ \s]*본[ \s]*총[ \s]*계")

	re_asset_list.append(re_asset_current)
	re_asset_list.append(re_asset_current_sub1)
	re_asset_list.append(re_asset_current_sub2)		
	re_asset_list.append(re_asset_current_sub3)		
	re_asset_list.append(re_asset_non_current)
	re_asset_list.append(re_asset_non_current_sub1)	
	re_asset_list.append(re_asset_non_current_sub2)	
	re_asset_list.append(re_asset_sum)
	re_asset_list.append(re_liability_current)
	re_asset_list.append(re_liability_current_sub1)
	re_asset_list.append(re_liability_current_sub2)		
	re_asset_list.append(re_liability_current_sub3)		
	re_asset_list.append(re_liability_non_current)
	re_asset_list.append(re_liability_non_current_sub1)	
	re_asset_list.append(re_liability_non_current_sub2)	
	re_asset_list.append(re_liability_non_current_sub3)	
	re_asset_list.append(re_liability_non_current_sub4)	
	re_asset_list.append(re_liability_sum)
	re_asset_list.append(re_equity)
	re_asset_list.append(re_equity_sub1)
	re_asset_list.append(re_equity_sub2)		
	re_asset_list.append(re_equity_sum)

	balance_sheet_sub_list = {}
	balance_sheet_sub_list["asset_current"]					=	0.0
	balance_sheet_sub_list["asset_current_sub1"]			=	0.0
	balance_sheet_sub_list["asset_current_sub2"]			=	0.0
	balance_sheet_sub_list["asset_current_sub3"]			=	0.0
	balance_sheet_sub_list["asset_non_current"]				=	0.0
	balance_sheet_sub_list["asset_non_current_sub1"]		=	0.0
	balance_sheet_sub_list["asset_non_current_sub2"]		=	0.0
	balance_sheet_sub_list["asset_sum"]						=	0.0
	balance_sheet_sub_list['year']							=	year
	balance_sheet_sub_list["liability_current"]				=	0.0
	balance_sheet_sub_list["liability_current_sub1"]		=	0.0
	balance_sheet_sub_list["liability_current_sub2"]		=	0.0
	balance_sheet_sub_list["liability_current_sub3"]		=	0.0
	balance_sheet_sub_list["liability_non_current"]			=	0.0
	balance_sheet_sub_list["liability_non_current_sub1"]	=	0.0
	balance_sheet_sub_list["liability_non_current_sub2"]	=	0.0
	balance_sheet_sub_list["liability_non_current_sub3"]	=	0.0
	balance_sheet_sub_list["liability_non_current_sub4"]	=	0.0
	balance_sheet_sub_list["liability_sum"]					=	0.0
	balance_sheet_sub_list["equity"]						=	0.0
	balance_sheet_sub_list["equity_sub1"]					=	0.0
	balance_sheet_sub_list["equity_sub2"]					=	0.0
	balance_sheet_sub_list["equity_sum"]					=	0.0

	balance_sheet_key_list = []
	
	balance_sheet_key_list.append("asset_current")
	balance_sheet_key_list.append("asset_current_sub1")
	balance_sheet_key_list.append("asset_current_sub2")
	balance_sheet_key_list.append("asset_current_sub3")
	balance_sheet_key_list.append("asset_non_current")
	balance_sheet_key_list.append("asset_non_current_sub1")
	balance_sheet_key_list.append("asset_non_current_sub2")
	balance_sheet_key_list.append("asset_sum")
	balance_sheet_key_list.append("liability_current")			
	balance_sheet_key_list.append("liability_current_sub1")		
	balance_sheet_key_list.append("liability_current_sub2")		
	balance_sheet_key_list.append("liability_current_sub3")		
	balance_sheet_key_list.append("liability_non_current")		
	balance_sheet_key_list.append("liability_non_current_sub1")	
	balance_sheet_key_list.append("liability_non_current_sub2")	
	balance_sheet_key_list.append("liability_non_current_sub3")	
	balance_sheet_key_list.append("liability_non_current_sub4")	
	balance_sheet_key_list.append("liability_sum")				
	balance_sheet_key_list.append("equity")						
	balance_sheet_key_list.append("equity_sub1")				
	balance_sheet_key_list.append("equity_sub2")				
	balance_sheet_key_list.append("equity_sum")					
	
	trs = balance_sheet_table.findAll("tr")

	# Balance sheet statement
	if (len(trs) != 2):
		for tr in trs:
			#print("trs", len(trs))
			tds = tr.findAll("td")
			#print("tds", len(tds))
			try:
				if (len(tds) != 0):
					#print(tds[0].text.strip())
					value = 0.0
					for i in range(len(re_asset_list)):
						if re_asset_list[i].search(tds[0].text.strip()):
							if len(tds)>4:
								if (tds[1].text.strip() != '') and (tds[1].text.strip() != '-'):
									value = find_value(tds[1].text.strip(), unit)
									break # for i in len(re_asset_list)
								elif (tds[2].text.strip() != '') and (tds[2].text.strip() != '-'):
									value = find_value(tds[2].text.strip(), unit)
									break # for i in len(re_asset_list)
							else:
								if (tds[1].text.strip() != '') and (tds[1].text.strip() != '-'):
									value = find_value(tds[1].text.strip(), unit)
									break # for i in len(re_asset_list)
					if value != 0.0:
						balance_sheet_sub_list[balance_sheet_key_list[i]] = value
			except Exception as e:
				print("NET INCOME PARSING ERROR in Balance sheet")
				print(e)
	# Special case
	## if (len(trs) != 2):
	else:	
		tr = trs[1]
		tds = tr.findAll("td")
		
		index_col = []
		prev = 0
		for a in tds[0].childGenerator():
			if (str(a) == "<br/>"):
				if (prev == 1):
					index_col.append('')	
				prev = 1
			else:
				prev = 0
				index_col.append(str(a).strip())	
		data_col = []
		prev = 0
		for b in tds[1].childGenerator():
			if (str(b) == "<br/>"):
				if (prev == 1):
					data_col.append('0')	
				prev = 1
			else:
				data_col.append(str(b))	
				prev = 0

		#print(index_col)
		#print(data_col)
		print(len(index_col))
		print(len(data_col))
		index_cnt = 0

		try:
			for (index) in (index_col):
				if (data_col[index_cnt].strip() != '') and (data_col[index_cnt].strip() != '-'):
					value = 0.0
					for i in range(len(re_asset_list)):
						if re_asset_list[i].search(index):
							value = find_value(data_col[index_cnt], unit)
							break
					if value != 0.0:
						balance_sheet_sub_list[balance_sheet_key_list[i]] = value
		except Exception as e:
			print("PARSING ERROR in BALANCE SHEET")
			print(e)

	print(balance_sheet_sub_list)
	return balance_sheet_sub_list


# Get information of cashflows statements
def scrape_cashflows(cashflow_table, year, unit):

	error_cashflows_list = []
	re_cashflow_list = []

	# Regular expression
	re_op_cashflow			= re.compile("((영업활동)|(영업활동으로[ \s]*인한)|(영업활동으로부터의))[ \s]*([순]*현금[ \s]*흐름)")
	re_op_cashflow_sub1 	= re.compile("((영업에서)|(영업으로부터))[ \s]*창출된[ \s]*현금(흐름)*")
	re_op_cashflow_sub2 	= re.compile("(연[ \s]*결[ \s]*)*당[ \s]*기[ \s]*순[ \s]*((이[ \s]*익)|(손[ \s]*익))")
	re_op_cashflow_sub3 	= re.compile("감[ \s]*가[ \s]*상[ \s]*각[ \s]*비")
	re_op_cashflow_sub4 	= re.compile("신[ \s]*탁[ \s]*계[ \s]*정[ \s]*대")
	
	re_invest_cashflow		= re.compile("투자[ \s]*활동[ \s]*현금[ \s]*흐름|투[ \s]*자[ \s]*활[ \s]*동[ \s]*으[ \s]*로[ \s]*인[ \s]*한[ \s]*[순]*현[ \s]*금[ \s]*흐[ \s]*름")
	re_invest_cashflow_sub1 = re.compile("유[ \s]*형[ \s]*자[ \s]*산[ \s]*의[ \s]*((취[ \s]*득)|(증[ \s]*가))")
	re_invest_cashflow_sub2 = re.compile("무[ \s]*형[ \s]*자[ \s]*산[ \s]*의[ \s]*((취[ \s]*득)|(증[ \s]*가))")
	re_invest_cashflow_sub3 = re.compile("토[ \s]*지[ \s]*의[ \s]*((취[ \s]*득)|(증[ \s]*가))")
	re_invest_cashflow_sub4 = re.compile("건[ \s]*물[ \s]*의[ \s]*((취[ \s]*득)|(증[ \s]*가))")
	re_invest_cashflow_sub5 = re.compile("구[ \s]*축[ \s]*물[ \s]*의[ \s]*((취[ \s]*득)|(증[ \s]*가))")
	re_invest_cashflow_sub6 = re.compile("기[ \s]*계[ \s]*장[ \s]*치[ \s]*의[ \s]*((취[ \s]*득)|(증[ \s]*가))")
	re_invest_cashflow_sub7 = re.compile("건[ \s]*설[ \s]*중[ \s]*인[ \s]*자[ \s]*산[ \s]*의[ \s]*((증[ \s]*가)|(취[ \s]*득))")
	re_invest_cashflow_sub8 = re.compile("차[ \s]*량[ \s]*운[ \s]*반[ \s]*구[ \s]*의[ \s]*((취[ \s]*득)|(증[ \s]*가))")
	re_invest_cashflow_sub9 = re.compile("비[ \s]*품[ \s]*의[ \s]*취[ \s]*득|비[ \s]*품[ \s]*의[ \s]*((증[ \s]*가)|(취[ \s]*득))")
	re_invest_cashflow_sub10= re.compile("공[ \s]*구[ \s]*기[ \s]*구[ \s]*의[ \s]*((취[ \s]*득)|(증[ \s]*가))")
	re_invest_cashflow_sub11= re.compile("시[ \s]*험[ \s]*연[ \s]*구[ \s]*설[ \s]*비[ \s]*의[ \s]*취[ \s]*득")
	re_invest_cashflow_sub12= re.compile("렌[ \s]*탈[ \s]*자[ \s]*산[ \s]*의[ \s]*((취[ \s]*득)|(증[ \s]*가))")
	re_invest_cashflow_sub13= re.compile("영[ \s]*업[ \s]*권[ \s]*의[ \s]*((취[ \s]*득)|(증[ \s]*가))")
	re_invest_cashflow_sub14= re.compile("산[ \s]*업[ \s]*재[ \s]*산[ \s]*권[ \s]*의[ \s]*((취[ \s]*득)|(증[ \s]*가))")
	re_invest_cashflow_sub15= re.compile("소[ \s]*프[ \s]*트[ \s]*웨[ \s]*어[ \s]*의[ \s]*((취[ \s]*득)|(증[ \s]*가))")
	re_invest_cashflow_sub16= re.compile("기[ \s]*타[ \s]*무[ \s]*형[ \s]*자[ \s]*산[ \s]*의[ \s]*((취[ \s]*득)|(증[ \s]*가))")
	re_invest_cashflow_sub17= re.compile("투[ \s]*자[ \s]*부[ \s]*통[ \s]*산[ \s]*의[ \s]*((취[ \s]*득)|(증[ \s]*가))")
	re_invest_cashflow_sub18= re.compile("관[ \s]*계[ \s]*기[ \s]*업[ \s]*투[ \s]*자[ \s]*의[ \s]*취[ \s]*득|관계[ \s]*기업[ \s]*투자[ \s]*주식의[ \s]*취득|지분법[ \s]*적용[ \s]*투자[ \s]*주식의[ \s]*취득")
	
	re_fin_cashflow			= re.compile("재무[ \s]*활동[ \s]*현금[ \s]*흐름|재무활동으로[ \s]*인한[ \s]*현금흐름")
	re_fin_cashflow_sub1	= re.compile("단기차입금의[ \s]*순증가")
	re_fin_cashflow_sub2	= re.compile("배당금[ \s]*지급|현금배당금의[ \s]*지급|배당금의[ \s]*지급|현금배당|보통주[ ]*배당[ ]*지급")
	re_fin_cashflow_sub3	= re.compile("자기주식의[ \s]*취득")
	re_start_cash			= re.compile("기초[ ]*현금[ ]*및[ ]*현금성[ ]*자산|기초의[ \s]*현금[ ]*및[ ]*현금성[ ]*자산|기[ \s]*초[ \s]*의[ \s]*현[ \s]*금|기[ \s]*초[ \s]*현[ \s]*금")
	re_end_cash				= re.compile("기말[ ]*현금[ ]*및[ ]*현금성[ ]*자산|기말의[ \s]*현금[ ]*및[ ]*현금성[ ]*자산|기[ \s]*말[ \s]*의[ \s]*현[ \s]*금|기[ \s]*말[ \s]*현[ \s]*금")


	re_cashflow_list.append(re_op_cashflow)
	re_cashflow_list.append(re_op_cashflow_sub1) 	
	re_cashflow_list.append(re_op_cashflow_sub2) 	
	re_cashflow_list.append(re_op_cashflow_sub3) 	
	re_cashflow_list.append(re_op_cashflow_sub4) 	
	
	re_cashflow_list.append(re_invest_cashflow)		
	re_cashflow_list.append(re_invest_cashflow_sub1) 
	re_cashflow_list.append(re_invest_cashflow_sub2) 
	re_cashflow_list.append(re_invest_cashflow_sub3) 
	re_cashflow_list.append(re_invest_cashflow_sub4) 
	re_cashflow_list.append(re_invest_cashflow_sub5) 
	re_cashflow_list.append(re_invest_cashflow_sub6) 
	re_cashflow_list.append(re_invest_cashflow_sub7) 
	re_cashflow_list.append(re_invest_cashflow_sub8) 
	re_cashflow_list.append(re_invest_cashflow_sub9) 
	re_cashflow_list.append(re_invest_cashflow_sub10)
	re_cashflow_list.append(re_invest_cashflow_sub11)
	re_cashflow_list.append(re_invest_cashflow_sub12)
	re_cashflow_list.append(re_invest_cashflow_sub13)
	re_cashflow_list.append(re_invest_cashflow_sub14)
	re_cashflow_list.append(re_invest_cashflow_sub15)
	re_cashflow_list.append(re_invest_cashflow_sub16)
	re_cashflow_list.append(re_invest_cashflow_sub17)
	re_cashflow_list.append(re_invest_cashflow_sub18)
	
	re_cashflow_list.append(re_fin_cashflow)		
	re_cashflow_list.append(re_fin_cashflow_sub1)	
	re_cashflow_list.append(re_fin_cashflow_sub2)	
	re_cashflow_list.append(re_fin_cashflow_sub3)	
	re_cashflow_list.append(re_start_cash)
	re_cashflow_list.append(re_end_cash)


	# 영업현금흐름
	## 영업에서 창출된 현금흐름
	## 당기순이익
	## 신탁계정대
	# 투자현금흐름
	## 유형자산의 취득
	## 무형자산의 취득
	## 토지의 취득
	## 건물의 취득
	## 구축물의 취득
	## 기계장치의 취득
	## 건설중인자산의증가
	## 차량운반구의 취득
	## 영업권의 취득
	## 산업재산권의 취득
	## 기타의무형자산의취득
	## 투자부동산의 취득
	## 관계기업투자의취득
	# 재무현금흐름
	## 단기차입금의 순증가
	## 배당금 지급
	## 자기주식의 취득
	# 기초 현금 및 현금성자산
	# 기말 현금 및 현금성자산

	cashflow_sub_list = {}
	
	cashflow_sub_list['year']					= year
	cashflow_sub_list["op_cashflow"]			= 0.0
	cashflow_sub_list["op_cashflow_sub1"]		= 0.0
	cashflow_sub_list["op_cashflow_sub2"]		= 0.0
	cashflow_sub_list["op_cashflow_sub3"]		= 0.0
	cashflow_sub_list["op_cashflow_sub4"]		= 0.0
	cashflow_sub_list["invest_cashflow"]		= 0.0
	cashflow_sub_list["invest_cashflow_sub1"]	= 0.0
	cashflow_sub_list["invest_cashflow_sub2"]	= 0.0
	cashflow_sub_list["invest_cashflow_sub3"]	= 0.0
	cashflow_sub_list["invest_cashflow_sub4"]	= 0.0
	cashflow_sub_list["invest_cashflow_sub5"]	= 0.0
	cashflow_sub_list["invest_cashflow_sub6"]	= 0.0
	cashflow_sub_list["invest_cashflow_sub7"]	= 0.0
	cashflow_sub_list["invest_cashflow_sub8"]	= 0.0
	cashflow_sub_list["invest_cashflow_sub9"]	= 0.0
	cashflow_sub_list["invest_cashflow_sub10"]	= 0.0
	cashflow_sub_list["invest_cashflow_sub11"]	= 0.0
	cashflow_sub_list["invest_cashflow_sub12"]	= 0.0
	cashflow_sub_list["invest_cashflow_sub13"]	= 0.0
	cashflow_sub_list["invest_cashflow_sub14"]	= 0.0
	cashflow_sub_list["invest_cashflow_sub15"]	= 0.0
	cashflow_sub_list["invest_cashflow_sub16"]	= 0.0
	cashflow_sub_list["invest_cashflow_sub17"]	= 0.0
	cashflow_sub_list["invest_cashflow_sub18"]	= 0.0
	cashflow_sub_list["fin_cashflow"]			= 0.0
	cashflow_sub_list["fin_cashflow_sub1"]		= 0.0
	cashflow_sub_list["fin_cashflow_sub2"]		= 0.0
	cashflow_sub_list["fin_cashflow_sub3"]		= 0.0
	cashflow_sub_list["start_cash"]				= 0.0
	cashflow_sub_list["end_cash"]				= 0.0

	cashflow_key_list = []

	cashflow_key_list.append("op_cashflow")
	cashflow_key_list.append("op_cashflow_sub1")
	cashflow_key_list.append("op_cashflow_sub2")
	cashflow_key_list.append("op_cashflow_sub3")
	cashflow_key_list.append("op_cashflow_sub4")
	cashflow_key_list.append("invest_cashflow")
	cashflow_key_list.append("invest_cashflow_sub1")
	cashflow_key_list.append("invest_cashflow_sub2")
	cashflow_key_list.append("invest_cashflow_sub3")
	cashflow_key_list.append("invest_cashflow_sub4")
	cashflow_key_list.append("invest_cashflow_sub5")
	cashflow_key_list.append("invest_cashflow_sub6")
	cashflow_key_list.append("invest_cashflow_sub7")
	cashflow_key_list.append("invest_cashflow_sub8")
	cashflow_key_list.append("invest_cashflow_sub9")
	cashflow_key_list.append("invest_cashflow_sub10")
	cashflow_key_list.append("invest_cashflow_sub11")
	cashflow_key_list.append("invest_cashflow_sub12")
	cashflow_key_list.append("invest_cashflow_sub13")
	cashflow_key_list.append("invest_cashflow_sub14")
	cashflow_key_list.append("invest_cashflow_sub15")
	cashflow_key_list.append("invest_cashflow_sub16")
	cashflow_key_list.append("invest_cashflow_sub17")
	cashflow_key_list.append("invest_cashflow_sub18")
	cashflow_key_list.append("fin_cashflow")
	cashflow_key_list.append("fin_cashflow_sub1")
	cashflow_key_list.append("fin_cashflow_sub2")
	cashflow_key_list.append("fin_cashflow_sub3")
	cashflow_key_list.append("start_cash")
	cashflow_key_list.append("end_cash")

	#net_income = 0.0
	
	#print("len(trs)", len(trs))
	
	trs = cashflow_table.findAll("tr")
			
	# CASHFLOW statement
	if (len(trs) != 2):
		for tr in trs:
			#print("trs", len(trs))
			tds = tr.findAll("td")
			#print("tds", len(tds))
			try:
				if (len(tds) != 0):
					#print(tds[0].text.strip())

					value = 0.0
					for i in range(len(re_cashflow_list)):
						if re_cashflow_list[i].search(tds[0].text.strip()):
							if len(tds)>4:
								if (tds[1].text.strip() != '') and (tds[1].text.strip() != '-'):
									value = find_value(tds[1].text.strip(), unit)
									break # for i in len(re_cashflow_list)
								elif (tds[2].text.strip() != '') and (tds[2].text.strip() != '-'):
									value = find_value(tds[2].text.strip(), unit)
									break # for i in len(re_cashflow_list)
							else:
								if (tds[1].text.strip() != '') and (tds[1].text.strip() != '-'):
									value = find_value(tds[1].text.strip(), unit)
									break # for i in len(re_cashflow_list)
					if value != 0.0:
						cashflow_sub_list[cashflow_key_list[i]] = value
					# No matching case
					else:
						error_cashflows_list.append(tds[0].text.strip())
			except Exception as e:
				print("NET INCOME PARSING ERROR in Cashflows")
				cashflow_sub_list["op_cashflow_sub1"] = "PARSING ERROR"
				print(e)
	# Special case
	## if (len(trs) != 2):
	else:	
		tr = trs[1]
		tds = tr.findAll("td")
		
		index_col = []
		prev = 0
		for a in tds[0].childGenerator():
			if (str(a) == "<br/>"):
				if (prev == 1):
					index_col.append('')	
				prev = 1
			else:
				prev = 0
				index_col.append(str(a).strip())	
		data_col = []
		prev = 0
		for b in tds[1].childGenerator():
			if (str(b) == "<br/>"):
				if (prev == 1):
					data_col.append('0')	
				prev = 1
			else:
				data_col.append(str(b))	
				prev = 0

		#print(index_col)
		#print(data_col)
		print(len(index_col))
		print(len(data_col))
		index_cnt = 0

		try:
			for (index) in (index_col):
				if (data_col[index_cnt].strip() != '') and (data_col[index_cnt].strip() != '-'):
					value = 0.0
					for i in range(len(re_asset_list)):
						if re_cashflow_list[i].search(index):
							value = find_value(data_col[index_cnt], unit)
							break
					if value != 0.0:
						cashflow_sub_list[cashflow_key_list[i]] = value
		except Exception as e:
			print("PARSING ERROR")
			cashflow_sub_list["op_cashflow_sub1"] = "PARSING ERROR"
			print(e)

	print(cashflow_sub_list)
	print(error_cashflows_list)
	return cashflow_sub_list

# Get information of income statements
def scrape_income_statement(income_table, year, unit):

	#매출액
	#매출원가
	#매출총이익
	#판매비와관리비
	#영업이익
	#기타수익
	#기타비용
	#금융수익
	#금융비용
	#법인세비용차감전순이익
	#번인세비용
	#당기순이익
	#기본주당이익

	re_income_list = []
	
	# Regular expression
	re_sales			=	re.compile("매[ \s]*출[ \s]*액")
	re_sales_sub1		= 	re.compile("매[ \s]*출[ \s]*원[ \s]*가")
	re_sales_sub2		= 	re.compile("매[ \s]*출[ \s]*총[ \s]*이[ \s]*익")
	re_sales_sub3		= 	re.compile("판[ \s]*매[ \s]*비[ \s]*와[ \s]*관[ \s]*리[ \s]*비")
	re_op_income		= 	re.compile("영[ \s]*업[ \s]*이[ \s]*익")
	re_op_income_sub1	= 	re.compile("기[ \s]*타[ \s]*수[ \s]*익")
	re_op_income_sub2	= 	re.compile("기[ \s]*타[ \s]*비[ \s]*용")
	re_op_income_sub3	= 	re.compile("금[ \s]*융[ \s]*수[ \s]*익")
	re_op_income_sub4	= 	re.compile("금[ \s]*융[ \s]*비[ \s]*용")
	re_op_income_sub6	= 	re.compile("영[ \s]*업[ \s]*외[ \s]*수[ \s]*익")
	re_op_income_sub7	= 	re.compile("영[ \s]*업[ \s]*외[ \s]*비[ \s]*용")
	re_op_income_sub5	= 	re.compile("법[ \s]*인[ \s]*세[ \s]*비[ \s]*용[ \s]*차[ \s]*감[ \s]*전[ \s]*순[ \s]*이[ \s]*익|법[ \s]*인[ \s]*세[ \s]*차[ \s]*감[ \s]*전[ \s]*계[ \s]*속[ \s]*영[ \s]*업[ \s]*순[ \s]*이[ \s]*익|법인세[ \s]*차감전[ \s]*순이익|법인세차감전계속영업이익|법인세비용차감전이익|법인세비용차감전계속영업[순]*이익|법인세비용차감전당기순이익|법인세비용차감전순이익|법인세비용차감전[ \s]*계속사업이익|법인세비용차감전순손익")
	re_tax				=	re.compile("법[ \s]*인[ \s]*세[ \s]*비[ \s]*용")
	re_net_income		=	re.compile("^순[ \s]*이[ \s]*익|^당[ \s]*기[ \s]*순[ \s]*이[ \s]*익|^연[ ]*결[ ]*당[ ]*기[ ]*순[ ]*이[ ]*익|지배기업의 소유주에게 귀속되는 당기순이익|분기순이익|당\(분\)기순이익|\.[ \s]*당[ \s]*기[ \s]*순[ \s]*이[ \s]*익|당분기연결순이익")
	re_eps				=	re.compile("기[ \s]*본[ \s]*주[ \s]*당[ \s]*((수[ \s]*익)|(순[ \s]*이[ \s]*익))")

	re_income_list.append(re_sales)	
	re_income_list.append(re_sales_sub1)		 	
	re_income_list.append(re_sales_sub2)		 	
	re_income_list.append(re_sales_sub3)		 	
	re_income_list.append(re_op_income)		 	
	re_income_list.append(re_op_income_sub1)	 	
	re_income_list.append(re_op_income_sub2)	 	
	re_income_list.append(re_op_income_sub3)	 	
	re_income_list.append(re_op_income_sub4)	 	
	re_income_list.append(re_op_income_sub5)	 	
	re_income_list.append(re_op_income_sub6)	 	
	re_income_list.append(re_op_income_sub7)	 	
	re_income_list.append(re_tax)
	re_income_list.append(re_net_income)
	re_income_list.append(re_eps)				

	income_statement_sub_list = {}
	income_statement_sub_list["sales"]				=	0.0
	income_statement_sub_list["sales_sub1"]			=	0.0
	income_statement_sub_list["sales_sub2"]			=	0.0
	income_statement_sub_list["sales_sub3"]			=	0.0
	income_statement_sub_list["op_income"]		 	=	0.0
	income_statement_sub_list["op_income_sub1"]		=	0.0
	income_statement_sub_list["op_income_sub2"]		=	0.0
	income_statement_sub_list["op_income_sub3"]		=	0.0
	income_statement_sub_list["op_income_sub4"]		=	0.0
	income_statement_sub_list["op_income_sub5"]		=	0.0
	income_statement_sub_list["op_income_sub6"]		=	0.0
	income_statement_sub_list["op_income_sub7"]		=	0.0
	income_statement_sub_list["tax"]				=	0.0
	income_statement_sub_list["net_income"]			=	0.0
	income_statement_sub_list["eps"]				=	0.0
	income_statement_sub_list['year']				=	year

	income_statement_key_list = []
	income_statement_key_list.append("sales")			
	income_statement_key_list.append("sales_sub1")		
	income_statement_key_list.append("sales_sub2")		
	income_statement_key_list.append("sales_sub3")		
	income_statement_key_list.append("op_income")		
	income_statement_key_list.append("op_income_sub1")	
	income_statement_key_list.append("op_income_sub2")	
	income_statement_key_list.append("op_income_sub3")	
	income_statement_key_list.append("op_income_sub4")	
	income_statement_key_list.append("op_income_sub5")	
	income_statement_key_list.append("op_income_sub6")	
	income_statement_key_list.append("op_income_sub7")	
	income_statement_key_list.append("tax")			
	income_statement_key_list.append("net_income")		
	income_statement_key_list.append("eps")			

	trs = income_table.findAll("tr")

	# Income statement
	if (len(trs) != 2):
		for income_tr in trs:
			tds = income_tr.findAll("td")
			try:
				if (len(tds) != 0):
					#print(tds[0].text.strip())
					value = 0.0
					for i in range(len(re_income_list)):
						if re_income_list[i].search(tds[0].text.strip()):
							if len(tds)>4:
								if (tds[1].text.strip() != '') and (tds[1].text.strip() != '-'):
									value = find_value(tds[1].text.strip(), unit)
									break # for i in len(re_income_list)
								elif (tds[2].text.strip() != '') and (tds[2].text.strip() != '-'):
									value = find_value(tds[2].text.strip(), unit)
									break # for i in len(re_income_list)
							else:
								if (tds[1].text.strip() != '') and (tds[1].text.strip() != '-'):
									value = find_value(tds[1].text.strip(), unit)
									break # for i in len(re_income_list)
					if value != 0.0:
						income_statement_sub_list[income_statement_key_list[i]] = value
			except Exception as e:
				print("NET INCOME PARSING ERROR in Income statement")
				print(e)
				net_income = 0.0
	## if (len(trs) != 2):
	else:	
		income_tr = trs[1]
		tds = income_tr.findAll("td")
		
		index_col = []
		prev = 0
		for a in tds[0].childGenerator():
			if (str(a) == "<br/>"):
				if (prev == 1):
					index_col.append('')	
				prev = 1
			else:
				prev = 0
				index_col.append(str(a).strip())	
		data_col = []
		prev = 0
		for b in tds[1].childGenerator():
			if (str(b) == "<br/>"):
				if (prev == 1):
					data_col.append('0')	
				prev = 1
			else:
				data_col.append(str(b))	
				prev = 0
		
		print(len(index_col))
		print(len(data_col))
		index_cnt = 0

		try:
			for (index) in (index_col):
				if (data_col[index_cnt].strip() != '') and (data_col[index_cnt].strip() != '-'):
					value = 0.0
					for i in range(len(re_income_list)):
						if re_income_list[i].search(index):
							value = find_value(data_col[index_cnt], unit)
							break
					if value != 0.0:
						balance_sheet_sub_list[income_statement_key_list[i]] = value
		except Exception as e:
			print("PARSING ERROR in INCOME STATEMENT")
			print(e)

	print(income_statement_sub_list)
	return income_statement_sub_list

# Main function
def main():

	# Default
	#corp = "삼성전자"
	corp = "LG화학"
	workbook_name = "DART_financial_statement.xlsx"

	try:
		#opts, args = getopt.getopt(sys.argv[1:], "m:s:e:c:o:h", ["mode=", "start=", "end=", "corp=", "output", "help"])
		opts, args = getopt.getopt(sys.argv[1:], "c:o:h", ["corp=", "output", "help"])
	except getopt.GetoptError as err:
		print(err)
		sys.exit(2)
	for option, argument in opts:
		if option == "-h" or option == "--help":
			help_msg = """
================================================================================
-c or --corp <name>     :  Corporation name
-o or --output <name>	:  Output file name
-h or --help            :  Show help messages

<Example>
>> python dart_dividends.py -m 0 -s 20171115 -e 20171215 -o out_file_name
>> python dart_dividends.py -m 1 -c S-Oil
================================================================================
					"""
			print(help_msg)
			sys.exit(2)
		elif option == "--corp" or option == "-c":
			corp = argument
		elif option == "--output" or option == "-o":
			workbook_name = argument + ".xlsx"

	re_income_find = re.compile("법[ \s]*인[ \s]*세[ \s]*비[ \s]*용[ \s]*차[ \s]*감[ \s]*전[ \s]*순[ \s]*이[ \s]*익|법[ \s]*인[ \s]*세[ \s]*차[ \s]*감[ \s]*전[ \s]*계[ \s]*속[ \s]*영[ \s]*업[ \s]*순[ \s]*이[ \s]*익|법인세[ \s]*차감전[ \s]*순이익|법인세차감전계속영업이익|법인세비용차감전이익|법인세비용차감전계속영업[순]*이익|법인세비용차감전당기순이익|법인세비용차감전순이익|법인세비용차감전[ \s]*계속사업이익|법인세비용차감전순손익")
	re_cashflow_find = re.compile("영업활동[ \s]*현금[ \s]*흐름|영업활동으로[ \s]*인한[ \s]*[순]*현금[ \s]*흐름|영업활동으로부터의[ \s]*현금흐름")
	#re_balance_sheet_find = re.compile("^[ \s]*유[ \s]*동[ \s]*자[ \s]*산|\.[ \s]*유[ \s]*동[ \s]*자[ \s]*산")
	re_balance_sheet_find = re.compile("현[ \s]*금[ \s]*및[ \s]*현[ \s]*금[ \s]*((성[ \s]*자[ \s]*산)|(등[ \s]*가[ \s]*물))")

	### PART I - Read Excel file for stock lists
	num_stock = 2040
	#num_stock = 100
	input_file = "basic_20171221.xlsx"
	cur_dir = os.getcwd()
	workbook_read_name = input_file
	
	stock_cat_list = []
	stock_name_list = []
	stock_num_list = []
	stock_url_list = []
	
	workbook_read = xlrd.open_workbook(os.path.join(cur_dir, workbook_read_name))
	sheet_list = workbook_read.sheets()
	sheet1 = sheet_list[0]

	for i in range(num_stock):
		stock_cat_list.append(sheet1.cell(i+1,0).value)
		stock_name_list.append(sheet1.cell(i+1,1).value)
		stock_num_list.append(sheet1.cell(i+1,2).value)
		stock_url_list.append(sheet1.cell(i+1,3).value)

	find_index = stock_name_list.index(corp)

	stock_code = ""

	if find_index != -1:
		stock_code = stock_num_list[find_index]
		stock_cat = stock_cat_list[find_index]
	else:
		print("STOCK CODE ERROR")

	# URL
	#url_templete = "http://dart.fss.or.kr/dsab002/search.ax?reportName=%s&&maxResults=100&&textCrpNm=%s"
	url_templete = "http://dart.fss.or.kr/dsab002/search.ax?reportName=%s&&maxResults=100&&textCrpNm=%s&&startDate=%s&&endDate=%s"
	headers = {'Cookie':'DSAB002_MAXRESULTS=5000;'}
	
	dart_div_list = []
	cashflow_list = []
	
	year = 2017
	start_day = datetime(2007,1,1)
	#start_day = datetime(2000,1,1)
	#end_day = datetime(2002,11,15)
	end_day = datetime(2017,11,15)
	delta = end_day - start_day

	# 사업보고서
	report = "%EC%82%AC%EC%97%85%EB%B3%B4%EA%B3%A0%EC%84%9C"
	# 분기보고서
	report2 = "%EB%B6%84%EA%B8%B0%EB%B3%B4%EA%B3%A0%EC%84%9C" 
	start_day2 = datetime(2017,10,15)
	end_day2 = datetime(2017,11,17)


	# 최신 분기보고서 읽기
	handle = urllib.request.urlopen(url_templete % (report2, urllib.parse.quote(corp), start_day2.strftime('%Y%m%d'), end_day2.strftime('%Y%m%d')))

	data = handle.read()
	soup = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')
	
	table = soup.find('table')
	trs = table.findAll('tr')
	tds = table.findAll('td')
	counts = len(tds)
	
	if counts > 2:
		# Delay operation
		#time.sleep(20)
	
		link_list = []
		date_list = []
		corp_list = []
		market_list = []
		title_list = []
		reporter_list = []
		cashflow_list = []
		balance_sheet_list = []
		income_statement_list = []

		# recent report
		tr = trs[1]
		time.sleep(2)
		tds = tr.findAll('td')
		link = 'http://dart.fss.or.kr' + tds[2].a['href']
		date = tds[4].text.strip().replace('.', '-')
		corp_name = tds[1].text.strip()
		market = tds[1].img['title']
		title = " ".join(tds[2].text.split())
		reporter = tds[3].text.strip()

		link_list.append(link)
		date_list.append(date)
		corp_list.append(corp_name)
		market_list.append(market)
		title_list.append(title)
		reporter_list.append(reporter)
	
		dart_div_sublist = []

		year = int(date[0:4])
		print(corp_name)
		print(title)
		print(date)
		handle = urllib.request.urlopen(link)
		data = handle.read()
		soup2 = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')
		print(link)
		
		head_lines = soup2.find('head').text.split("\n")
		#print(head_lines)
		re_tree_find1 = re.compile("2.[ ]*연결재무제표")
		re_tree_find1_bak = re.compile("4.[ ]*재무제표")
		line_num = 0
		line_find = 0
		for head_line in head_lines:
			#print(head_line)
			if (re_tree_find1.search(head_line)):
				line_find = line_num
				break
			line_num = line_num + 1
		
		line_num = 0
		line_find_bak = 0
		for head_line in head_lines:
			if (re_tree_find1_bak.search(head_line)):
				line_find_bak = line_num
				break
			line_num = line_num + 1


		if(line_find != 0):
		
			line_words = head_lines[line_find+4].split("'")
			#print(line_words)
			rcpNo = line_words[1]
			dcmNo = line_words[3]
			eleId = line_words[5]
			offset = line_words[7]
			length = line_words[9]

			dart = soup2.find_all(string=re.compile('dart.dtd'))
			dart2 = soup2.find_all(string=re.compile('dart2.dtd'))
			dart3 = soup2.find_all(string=re.compile('dart3.xsd'))

			if len(dart3) != 0:
				link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=" + eleId + "&offset=" + offset + "&length=" + length + "&dtd=dart3.xsd"
			elif len(dart2) != 0:
				link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=" + eleId + "&offset=" + offset + "&length=" + length + "&dtd=dart2.dtd"
			elif len(dart) != 0:
				link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=" + eleId + "&offset=" + offset + "&length=" + length + "&dtd=dart.dtd"
			else:
				link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=0&offset=0&length=0&dtd=HTML"  
			
			print(link2)

			#try:
			handle = urllib.request.urlopen(link2)
			print(handle)
			data = handle.read()
			soup3 = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')

			tables = soup3.findAll("table")

			# 2. 연결재무제표가 비어 있는 경우
			if (len(tables) == 0):
				line_words = head_lines[line_find_bak+4].split("'")
				#print(line_words)
				rcpNo = line_words[1]
				dcmNo = line_words[3]
				eleId = line_words[5]
				offset = line_words[7]
				length = line_words[9]

				dart = soup2.find_all(string=re.compile('dart.dtd'))
				dart2 = soup2.find_all(string=re.compile('dart2.dtd'))
				dart3 = soup2.find_all(string=re.compile('dart3.xsd'))

				if len(dart3) != 0:
					link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=" + eleId + "&offset=" + offset + "&length=" + length + "&dtd=dart3.xsd"
				elif len(dart2) != 0:
					link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=" + eleId + "&offset=" + offset + "&length=" + length + "&dtd=dart2.dtd"
				elif len(dart) != 0:
					link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=" + eleId + "&offset=" + offset + "&length=" + length + "&dtd=dart.dtd"
				else:
					link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=0&offset=0&length=0&dtd=HTML"  
				
				print(link2)
				
				handle = urllib.request.urlopen(link2)
				print(handle)
				data = handle.read()
				soup3 = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')
				tables = soup3.findAll("table")

			cnt = 0
			table_num = 0

			for table in tables:
				if (re_cashflow_find.search(table.text)):
					table_num = cnt
					break
				cnt = cnt + 1
			
			print("table_num", table_num, "Tables", len(tables))
			cashflow_table = soup3.findAll("table")[table_num]
			
			cnt = 0
			table_income_num = 0
			for table in tables:
				if (re_income_find.search(table.text)):
					table_income_num = cnt
					break
				cnt = cnt + 1
			income_table = soup3.findAll("table")[table_income_num]
			#print("table_income_num", table_income_num, "Tables", len(tables))
			
			cnt = 0
			table_balance_num = 0
			for table in tables:
				if (re_balance_sheet_find.search(table.text)):
					table_balance_num = cnt
					break
				cnt = cnt + 1
			balance_table = soup3.findAll("table")[table_balance_num]
			print("table_balance_num", table_balance_num, "Tables", len(tables))
			
			unit = 100.0
			unit_find = 0
			re_unit1 = re.compile('단위[ \s]*:[ \s]*원')
			re_unit2 = re.compile('단위[ \s]*:[ \s]*백만원')
			re_unit3 = re.compile('단위[ \s]*:[ \s]*천원')

			# 원
			if len(soup3.findAll("table")[table_num-1](string=re_unit1)) != 0:
				unit = 100000000.0
				unit_find = 1
				#print("Unit ###1")
			# 백만원
			elif len(soup3.findAll("table")[table_num-1](string=re_unit2)) != 0:
				unit = 100.0
				unit_find = 1
				#print("Unit ###2")
			elif len(soup3.findAll("table")[table_num-1](string=re_unit3)) != 0:
				unit = 100000.0
				unit_find = 1
				#print("Unit ###3")

			if unit_find == 0:
				print ("UNIT NOT FOUND")
				if len(soup3.findAll(string=re_unit1)) != 0:
					unit = 100000000.0
				elif len(soup3.findAll(string=re_unit2)) != 0:
					unit = 100.0
				elif len(soup3.findAll(string=re_unit3)) != 0:
					unit = 100000.0
			
			cashflow_sub_list = scrape_cashflows(cashflow_table, 2017, unit)
			income_statement_sub_list = scrape_income_statement(income_table, 2017, unit)
			balance_sheet_sub_list = scrape_balance_sheet(balance_table, 2017, unit)
			
			cashflow_sub_list['net_income'] = income_statement_sub_list['net_income']

		## if(line_find != 0):
		else:
			print("FINDING LINE NUMBER ERROR")
			cashflow_sub_list = {}
			op_cashflow = 0.0
			op_cashflow_sub1 = "FINDING LINE NUMBER ERROR"
			op_cashflow_sub2 = 0.0
			invest_cashflow = 0.0
			invest_cashflow_sub1 = 0.0
			invest_cashflow_sub2 = 0.0
			invest_cashflow_sub3 = 0.0
			invest_cashflow_sub4 = 0.0
			invest_cashflow_sub5 = 0.0
			invest_cashflow_sub6 = 0.0
			invest_cashflow_sub7 = 0.0
			invest_cashflow_sub8 = 0.0
			invest_cashflow_sub9 = 0.0
			invest_cashflow_sub10 = 0.0
			invest_cashflow_sub11 = 0.0
			invest_cashflow_sub12 = 0.0
			invest_cashflow_sub13 = 0.0
			invest_cashflow_sub14 = 0.0
			invest_cashflow_sub15 = 0.0
			invest_cashflow_sub16 = 0.0
			invest_cashflow_sub17 = 0.0
			invest_cashflow_sub18 = 0.0
			fin_cashflow = 0.0
			fin_cashflow_sub1 = 0.0
			fin_cashflow_sub2 = 0.0
			fin_cashflow_sub3 = 0.0
			start_cash = 0.0
			end_cash = 0.0
			net_income = 0.0
			
			cashflow_sub_list['year'] = 2017
			cashflow_sub_list['op_cashflow'] = op_cashflow
			cashflow_sub_list['op_cashflow_sub1'] = op_cashflow_sub1
			cashflow_sub_list['op_cashflow_sub2'] = op_cashflow_sub2

			cashflow_sub_list['invest_cashflow'] = invest_cashflow
			cashflow_sub_list['invest_cashflow_sub1'] = invest_cashflow_sub1
			cashflow_sub_list['invest_cashflow_sub2'] = invest_cashflow_sub2
			cashflow_sub_list['invest_cashflow_sub3'] = invest_cashflow_sub3
			cashflow_sub_list['invest_cashflow_sub4'] = invest_cashflow_sub4
			cashflow_sub_list['invest_cashflow_sub5'] = invest_cashflow_sub5
			cashflow_sub_list['invest_cashflow_sub6'] = invest_cashflow_sub6
			cashflow_sub_list['invest_cashflow_sub7'] = invest_cashflow_sub7
			cashflow_sub_list['invest_cashflow_sub8'] = invest_cashflow_sub8
			cashflow_sub_list['invest_cashflow_sub9'] = invest_cashflow_sub9
			cashflow_sub_list['invest_cashflow_sub10'] = invest_cashflow_sub10
			cashflow_sub_list['invest_cashflow_sub11'] = invest_cashflow_sub11
			cashflow_sub_list['invest_cashflow_sub12'] = invest_cashflow_sub12
			cashflow_sub_list['invest_cashflow_sub13'] = invest_cashflow_sub13
			cashflow_sub_list['invest_cashflow_sub14'] = invest_cashflow_sub14
			cashflow_sub_list['invest_cashflow_sub15'] = invest_cashflow_sub15
			cashflow_sub_list['invest_cashflow_sub16'] = invest_cashflow_sub16
			cashflow_sub_list['invest_cashflow_sub17'] = invest_cashflow_sub17
			cashflow_sub_list['invest_cashflow_sub18'] = invest_cashflow_sub18
			
			cashflow_sub_list['fin_cashflow'] = fin_cashflow
			cashflow_sub_list['fin_cashflow_sub1'] = fin_cashflow_sub1
			cashflow_sub_list['fin_cashflow_sub2'] = fin_cashflow_sub2
			cashflow_sub_list['fin_cashflow_sub3'] = fin_cashflow_sub3

			cashflow_sub_list['start_cash'] = start_cash
			cashflow_sub_list['end_cash'] = end_cash
			
			cashflow_sub_list['net_income'] = net_income
			
			print(cashflow_sub_list)

			balance_sheet_sub_list = {}
			balance_sheet_sub_list["asset_current"]				=	0.0
			balance_sheet_sub_list["asset_current_sub1"]		=	0.0
			balance_sheet_sub_list["asset_current_sub2"]		=	0.0
			balance_sheet_sub_list["asset_current_sub3"]		=	0.0
			balance_sheet_sub_list["asset_non_current"]			=	0.0
			balance_sheet_sub_list["asset_non_current_sub1"]	=	0.0
			balance_sheet_sub_list["asset_non_current_sub2"]	=	0.0
			balance_sheet_sub_list["asset_sum"]					=	0.0
			balance_sheet_sub_list["liability_current"]				=	0.0
			balance_sheet_sub_list["liability_current_sub1"]		=	0.0
			balance_sheet_sub_list["liability_current_sub2"]		=	0.0
			balance_sheet_sub_list["liability_current_sub3"]		=	0.0
			balance_sheet_sub_list["liability_non_current"]			=	0.0
			balance_sheet_sub_list["liability_non_current_sub1"]	=	0.0
			balance_sheet_sub_list["liability_non_current_sub2"]	=	0.0
			balance_sheet_sub_list["liability_non_current_sub3"]	=	0.0
			balance_sheet_sub_list["liability_non_current_sub4"]	=	0.0
			balance_sheet_sub_list["liability_sum"]					=	0.0
			balance_sheet_sub_list["equity"]						=	0.0
			balance_sheet_sub_list["equity_sub1"]					=	0.0
			balance_sheet_sub_list["equity_sub2"]					=	0.0
			balance_sheet_sub_list["equity_sum"]					=	0.0
			balance_sheet_sub_list['year']						=	year-1
					
			income_statement_sub_list = {}
			income_statement_sub_list["sales"]				=	0.0
			income_statement_sub_list["sales_sub1"]			=	0.0
			income_statement_sub_list["sales_sub2"]			=	0.0
			income_statement_sub_list["sales_sub3"]			=	0.0
			income_statement_sub_list["op_income"]		 	=	0.0
			income_statement_sub_list["op_income_sub1"]		=	0.0
			income_statement_sub_list["op_income_sub2"]		=	0.0
			income_statement_sub_list["op_income_sub3"]		=	0.0
			income_statement_sub_list["op_income_sub4"]		=	0.0
			income_statement_sub_list["op_income_sub5"]		=	0.0
			income_statement_sub_list["op_income_sub6"]		=	0.0
			income_statement_sub_list["op_income_sub7"]		=	0.0
			income_statement_sub_list["tax"]				=	0.0
			income_statement_sub_list["net_income"]			=	0.0
			income_statement_sub_list["eps"]				=	0.0
			income_statement_sub_list['year']						=	year-1

			#except:
			#	print ("URL ERROR")
			
		dart_div_sublist.append(date)
		dart_div_sublist.append(corp_name)
		dart_div_sublist.append(market)
		dart_div_sublist.append(title)
		dart_div_sublist.append(link)
			
		dart_div_list.append(dart_div_sublist)
		cashflow_list.append(cashflow_sub_list)
		balance_sheet_list.append(balance_sheet_sub_list)
		income_statement_list.append(income_statement_sub_list)


	#handle = urllib.request.urlopen(url_templete % (report, urllib.parse.quote(corp)))
	#print("URL" + url_templete % (report, corp))
	handle = urllib.request.urlopen(url_templete % (report, urllib.parse.quote(corp), start_day.strftime('%Y%m%d'), end_day.strftime('%Y%m%d')))
	print("URL" + url_templete % (report, corp, start_day.strftime('%Y%m%d'), end_day.strftime('%Y%m%d')))

	data = handle.read()
	soup = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')
	
	table = soup.find('table')
	trs = table.findAll('tr')
	tds = table.findAll('td')
	counts = len(tds)
	#print(counts)

	#if counts > 0:
	if counts > 2:
		# Delay operation
		time.sleep(20)
	
		link_list = []
		date_list = []
		corp_list = []
		market_list = []
		title_list = []
		reporter_list = []
		tr_cnt = 0
		
		for tr in trs[1:]:
			tr_cnt = tr_cnt + 1
			time.sleep(2)
			tds = tr.findAll('td')
			link = 'http://dart.fss.or.kr' + tds[2].a['href']
			date = tds[4].text.strip().replace('.', '-')
			corp_name = tds[1].text.strip()
			market = tds[1].img['title']
			title = " ".join(tds[2].text.split())
			reporter = tds[3].text.strip()

			re_pass = re.compile("해외증권거래소등에신고한사업보고서등의국내신고")
			if (not re_pass.search(title)):
				link_list.append(link)
				date_list.append(date)
				corp_list.append(corp_name)
				market_list.append(market)
				title_list.append(title)
				reporter_list.append(reporter)

				dart_div_sublist = []

				year = int(date[0:4])
				print(corp_name)
				print(title)
				print(date)
				handle = urllib.request.urlopen(link)
				#print(link)
				data = handle.read()
				soup2 = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')
				#print(soup2)
				
				#print(type(soup2.find('head').text))
				head_lines = soup2.find('head').text.split("\n")
				#print(head_words)

				# From 2015 ~ now
				#if (year>2014):
				#	re_tree_find = re.compile("2. 연결재무제표")
				## From 2010 to 2014
				#elif (year>2009):
				#	re_tree_find = re.compile("재무제표 등")
				## From 2008 to 2009
				#elif (year>2007):
				#	re_tree_find = re.compile("1. 연결재무제표에 관한 사항")
				## From 2002 to 2007
				#elif (year>2001):
				#	re_tree_find = re.compile("4. 재무제표")
				#else:
				#	re_tree_find = re.compile("3. 재무제표")

				re_tree_find1 = re.compile("2. 연결재무제표")
				re_tree_find2 = re.compile("재무제표 등")
				re_tree_find3 = re.compile("1. 연결재무제표에 관한 사항")
				re_tree_find4 = re.compile("4. 재무제표")
				re_tree_find5 = re.compile("3. 재무제표")
				
				re_tree_find1_bak = re.compile("4.[ ]*재무제표")
				
				line_num = 0
				line_find = 0
				for head_line in head_lines:
					if (re_tree_find1.search(head_line)):
						line_find = line_num
						break
						#print(head_line)
					elif (re_tree_find2.search(head_line)):
						line_find = line_num
						break
					elif (re_tree_find3.search(head_line)):
						line_find = line_num
						break
					elif (re_tree_find4.search(head_line)):
						line_find = line_num
						break
					elif (re_tree_find5.search(head_line)):
						line_find = line_num
						break
					line_num = line_num + 1

				line_num = 0
				line_find_bak = 0
				for head_line in head_lines:
					if (re_tree_find1_bak.search(head_line)):
						line_find_bak = line_num
						break
					line_num = line_num + 1


				if(line_find != 0):
		
					#print(head_lines[line_find])
					#print(head_lines[line_find+1])
					#print(head_lines[line_find+2])
					#print(head_lines[line_find+3])
					#print(head_lines[line_find+4])

					line_words = head_lines[line_find+4].split("'")
					#print(line_words)
					rcpNo = line_words[1]
					dcmNo = line_words[3]
					eleId = line_words[5]
					offset = line_words[7]
					length = line_words[9]

					#test = soup2.find('a', {'href' : '#download'})['onclick']
					#words = test.split("'")
					#rcpNo = words[1]
					#dcmNo = words[3]
					
					dart = soup2.find_all(string=re.compile('dart.dtd'))
					dart2 = soup2.find_all(string=re.compile('dart2.dtd'))
					dart3 = soup2.find_all(string=re.compile('dart3.xsd'))

					if len(dart3) != 0:
						link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=" + eleId + "&offset=" + offset + "&length=" + length + "&dtd=dart3.xsd"
					elif len(dart2) != 0:
						link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=" + eleId + "&offset=" + offset + "&length=" + length + "&dtd=dart2.dtd"
					elif len(dart) != 0:
						link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=" + eleId + "&offset=" + offset + "&length=" + length + "&dtd=dart.dtd"
					else:
						link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=0&offset=0&length=0&dtd=HTML"  
					
					print(link2)

					#try:
					handle = urllib.request.urlopen(link2)
					#print(handle)
					data = handle.read()
					soup3 = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')
					#print(soup3)

					tables = soup3.findAll("table")
			
					# 2. 연결재무제표가 비어 있는 경우
					if (len(tables) == 0):
						line_words = head_lines[line_find_bak+4].split("'")
						#print(line_words)
						rcpNo = line_words[1]
						dcmNo = line_words[3]
						eleId = line_words[5]
						offset = line_words[7]
						length = line_words[9]

						dart = soup2.find_all(string=re.compile('dart.dtd'))
						dart2 = soup2.find_all(string=re.compile('dart2.dtd'))
						dart3 = soup2.find_all(string=re.compile('dart3.xsd'))

						if len(dart3) != 0:
							link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=" + eleId + "&offset=" + offset + "&length=" + length + "&dtd=dart3.xsd"
						elif len(dart2) != 0:
							link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=" + eleId + "&offset=" + offset + "&length=" + length + "&dtd=dart2.dtd"
						elif len(dart) != 0:
							link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=" + eleId + "&offset=" + offset + "&length=" + length + "&dtd=dart.dtd"
						else:
							link2 = "http://dart.fss.or.kr/report/viewer.do?rcpNo=" + rcpNo + "&dcmNo=" + dcmNo + "&eleId=0&offset=0&length=0&dtd=HTML"  
						
						print(link2)
						
						handle = urllib.request.urlopen(link2)
						print(handle)
						data = handle.read()
						soup3 = BeautifulSoup(data, 'html.parser', from_encoding='utf-8')
						tables = soup3.findAll("table")


					cnt = 0
					table_num = 0

					for table in tables:
						if (re_cashflow_find.search(table.text)):
							table_num = cnt
							break
						cnt = cnt + 1
					
					print("table_num", table_num, "Tables", len(tables))
					cashflow_table = soup3.findAll("table")[table_num]
					
					trs = cashflow_table.findAll("tr")
					
					cnt = 0
					table_income_num = 0
					for table in tables:
						if (re_income_find.search(table.text)):
							table_income_num = cnt
							break
						cnt = cnt + 1
					income_table = soup3.findAll("table")[table_income_num]
					print("table_income_num", table_income_num, "Tables", len(tables))
					
					cnt = 0
					table_balance_num = 0
					for table in tables:
						if (re_balance_sheet_find.search(table.text)):
							table_balance_num = cnt
							break
						cnt = cnt + 1
					balance_table = soup3.findAll("table")[table_balance_num]
					print("table_balance_num", table_balance_num, "Tables", len(tables))
			
					unit = 100.0
					unit_find = 0
					re_unit1 = re.compile('단위[ \s]*:[ \s]*원')
					re_unit2 = re.compile('단위[ \s]*:[ \s]*백만원')
					re_unit3 = re.compile('단위[ \s]*:[ \s]*천원')

					# 원
					if len(soup3.findAll("table")[table_num-1](string=re_unit1)) != 0:
						unit = 100000000.0
						unit_find = 1
						#print("Unit ###1")
					# 백만원
					elif len(soup3.findAll("table")[table_num-1](string=re_unit2)) != 0:
						unit = 100.0
						unit_find = 1
						#print("Unit ###2")
					elif len(soup3.findAll("table")[table_num-1](string=re_unit3)) != 0:
						unit = 100000.0
						unit_find = 1
						#print("Unit ###3")

					if unit_find == 0:
						print ("UNIT NOT FOUND")
						if len(soup3.findAll(string=re_unit1)) != 0:
							unit = 100000000.0
						elif len(soup3.findAll(string=re_unit2)) != 0:
							unit = 100.0
						elif len(soup3.findAll(string=re_unit3)) != 0:
							unit = 100000.0
			
					## 원
					#if len(soup3.findAll("table")[table_num-1](string=re.compile('단위[ ]*:[ ]*원'))) != 0:
					#	unit = 100000000.0
					## 백만원
					#elif len(soup3.findAll("table")[table_num-1](string=re.compile('단위[ ]*:[ ]*백만원'))) != 0:
					#	unit = 100.0
					#elif len(soup3.findAll("table")[table_num-1](string=re.compile('단위[ ]*:[ ]*천원'))) != 0:
					#	unit = 100000.0
				
					# Scrape data
					cashflow_sub_list = scrape_cashflows(cashflow_table, year-1, unit)
					income_statement_sub_list = scrape_income_statement(income_table, year-1, unit)
					balance_sheet_sub_list = scrape_balance_sheet(balance_table, year-1, unit)
					print(cashflow_sub_list)
					
					cashflow_sub_list['net_income'] = income_statement_sub_list['net_income']

				## if(line_find != 0):
				else:
					print("FINDING LINE NUMBER ERROR")
					cashflow_sub_list = {}
					op_cashflow = 0.0
					op_cashflow_sub1 = "FINDING LINE NUMBER ERROR"
					op_cashflow_sub2 = 0.0
					invest_cashflow = 0.0
					invest_cashflow_sub1 = 0.0
					invest_cashflow_sub2 = 0.0
					invest_cashflow_sub3 = 0.0
					invest_cashflow_sub4 = 0.0
					invest_cashflow_sub5 = 0.0
					invest_cashflow_sub6 = 0.0
					invest_cashflow_sub7 = 0.0
					invest_cashflow_sub8 = 0.0
					invest_cashflow_sub9 = 0.0
					invest_cashflow_sub10 = 0.0
					invest_cashflow_sub11 = 0.0
					invest_cashflow_sub12 = 0.0
					invest_cashflow_sub13 = 0.0
					invest_cashflow_sub14 = 0.0
					invest_cashflow_sub15 = 0.0
					invest_cashflow_sub16 = 0.0
					invest_cashflow_sub17 = 0.0
					invest_cashflow_sub18 = 0.0
					fin_cashflow = 0.0
					fin_cashflow_sub1 = 0.0
					fin_cashflow_sub2 = 0.0
					fin_cashflow_sub3 = 0.0
					start_cash = 0.0
					end_cash = 0.0
					net_income = 0.0
					
					cashflow_sub_list['year'] = year-1
					cashflow_sub_list['op_cashflow'] = op_cashflow
					cashflow_sub_list['op_cashflow_sub1'] = op_cashflow_sub1
					cashflow_sub_list['op_cashflow_sub2'] = op_cashflow_sub2

					cashflow_sub_list['invest_cashflow'] = invest_cashflow
					cashflow_sub_list['invest_cashflow_sub1'] = invest_cashflow_sub1
					cashflow_sub_list['invest_cashflow_sub2'] = invest_cashflow_sub2
					cashflow_sub_list['invest_cashflow_sub3'] = invest_cashflow_sub3
					cashflow_sub_list['invest_cashflow_sub4'] = invest_cashflow_sub4
					cashflow_sub_list['invest_cashflow_sub5'] = invest_cashflow_sub5
					cashflow_sub_list['invest_cashflow_sub6'] = invest_cashflow_sub6
					cashflow_sub_list['invest_cashflow_sub7'] = invest_cashflow_sub7
					cashflow_sub_list['invest_cashflow_sub8'] = invest_cashflow_sub8
					cashflow_sub_list['invest_cashflow_sub9'] = invest_cashflow_sub9
					cashflow_sub_list['invest_cashflow_sub10'] = invest_cashflow_sub10
					cashflow_sub_list['invest_cashflow_sub11'] = invest_cashflow_sub11
					cashflow_sub_list['invest_cashflow_sub12'] = invest_cashflow_sub12
					cashflow_sub_list['invest_cashflow_sub13'] = invest_cashflow_sub13
					cashflow_sub_list['invest_cashflow_sub14'] = invest_cashflow_sub14
					cashflow_sub_list['invest_cashflow_sub15'] = invest_cashflow_sub15
					cashflow_sub_list['invest_cashflow_sub16'] = invest_cashflow_sub16
					cashflow_sub_list['invest_cashflow_sub17'] = invest_cashflow_sub17
					cashflow_sub_list['invest_cashflow_sub18'] = invest_cashflow_sub18
					
					cashflow_sub_list['fin_cashflow'] = fin_cashflow
					cashflow_sub_list['fin_cashflow_sub1'] = fin_cashflow_sub1
					cashflow_sub_list['fin_cashflow_sub2'] = fin_cashflow_sub2
					cashflow_sub_list['fin_cashflow_sub3'] = fin_cashflow_sub3

					cashflow_sub_list['start_cash'] = start_cash
					cashflow_sub_list['end_cash'] = end_cash
					
					cashflow_sub_list['net_income'] = net_income
					
					print(cashflow_sub_list)

					balance_sheet_sub_list = {}
					balance_sheet_sub_list["asset_current"]				=	0.0
					balance_sheet_sub_list["asset_current_sub1"]		=	0.0
					balance_sheet_sub_list["asset_current_sub2"]		=	0.0
					balance_sheet_sub_list["asset_current_sub3"]		=	0.0
					balance_sheet_sub_list["asset_non_current"]			=	0.0
					balance_sheet_sub_list["asset_non_current_sub1"]	=	0.0
					balance_sheet_sub_list["asset_non_current_sub2"]	=	0.0
					balance_sheet_sub_list["asset_sum"]					=	0.0
					balance_sheet_sub_list["liability_current"]				=	0.0
					balance_sheet_sub_list["liability_current_sub1"]		=	0.0
					balance_sheet_sub_list["liability_current_sub2"]		=	0.0
					balance_sheet_sub_list["liability_current_sub3"]		=	0.0
					balance_sheet_sub_list["liability_non_current"]			=	0.0
					balance_sheet_sub_list["liability_non_current_sub1"]	=	0.0
					balance_sheet_sub_list["liability_non_current_sub2"]	=	0.0
					balance_sheet_sub_list["liability_non_current_sub3"]	=	0.0
					balance_sheet_sub_list["liability_non_current_sub4"]	=	0.0
					balance_sheet_sub_list["liability_sum"]					=	0.0
					balance_sheet_sub_list["equity"]						=	0.0
					balance_sheet_sub_list["equity_sub1"]					=	0.0
					balance_sheet_sub_list["equity_sub2"]					=	0.0
					balance_sheet_sub_list["equity_sum"]					=	0.0
					balance_sheet_sub_list['year']						=	year-1

					income_statement_sub_list = {}
					income_statement_sub_list["sales"]				=	0.0
					income_statement_sub_list["sales_sub1"]			=	0.0
					income_statement_sub_list["sales_sub2"]			=	0.0
					income_statement_sub_list["sales_sub3"]			=	0.0
					income_statement_sub_list["op_income"]		 	=	0.0
					income_statement_sub_list["op_income_sub1"]		=	0.0
					income_statement_sub_list["op_income_sub2"]		=	0.0
					income_statement_sub_list["op_income_sub3"]		=	0.0
					income_statement_sub_list["op_income_sub4"]		=	0.0
					income_statement_sub_list["op_income_sub5"]		=	0.0
					income_statement_sub_list["op_income_sub6"]		=	0.0
					income_statement_sub_list["op_income_sub7"]		=	0.0
					income_statement_sub_list["tax"]				=	0.0
					income_statement_sub_list["net_income"]			=	0.0
					income_statement_sub_list["eps"]				=	0.0

				#except:
				#	print ("URL ERROR")
				
				dart_div_sublist.append(date)
				dart_div_sublist.append(corp_name)
				dart_div_sublist.append(market)
				dart_div_sublist.append(title)
				dart_div_sublist.append(link)
				
				dart_div_list.append(dart_div_sublist)
				cashflow_list.append(cashflow_sub_list)
				balance_sheet_list.append(balance_sheet_sub_list)
				income_statement_list.append(income_statement_sub_list)

	write_excel_file(workbook_name, dart_div_list, cashflow_list, balance_sheet_list, income_statement_list, corp, stock_code, stock_cat)

# Main
if __name__ == "__main__":
	main()


