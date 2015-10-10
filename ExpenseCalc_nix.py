#!/usr/bin/python

# -------------------------------------------------------------------------------------------
# Developed By - Tom Thomas on 30 November 2012                                             |
# Current Version - 1.6                                                                     |
# Revision History                                                                          |
# 1.1 - 17 March 2013                                                                       |
#       Added functionality to allow tags have a special character(#) at the end to mean    |
#       that the amount should not added to the total expenditure but should be added to    |
#       the Tag/Category                                                                    |
# 1.2 - 13 April 2013                                                                       |
#       Added the functionality to get the TOTAL's column computing total of each category  |
#       as well as the total of header rows (Income, Expense, Savings).                     |
#       It also computes the percentage of each category against the Total Income in a new  |
#       column.                                                                             |
# 1.3 - 31 August 2014                                                                      |
#       Added the functionality to get iterate all input files if the sheet name does not   |
#       have 'NO' and give all the iterated output.                                         |
# 1.4 - 10 January 2015                                                                     |
#       Stripping all the input variable seperately into CONFIG file outside this script    |
#       Logging is introduced which writes to a log file in the DIR where this script runs  |
# 1.5 - 27 January 2015																		|
#		Fixed the Income Bug calling as Bug1self. 											|
# 1.6 - 29 January 2015																		|
#		Created HTML reporting based on Canvas.js showing current year Pie Chart and 		|
#       and multi-column bar chart to show trend.
# -------------------------------------------------------------------------------------------

import xlrd
import xlwt 

import datetime
import sys, os
import logging, subprocess

from itertools import groupby
from operator import itemgetter

# import matplotlib
# matplotlib.use('module://mplh5canvas.backend_h5canvas')
# from pylab import *
import webbrowser

# Config_ExpenCalc.py is a configuration file which should be in the current directory as this script
from Config_ExpCalc import * 

logging.basicConfig(filename='ExpnCalc.log', filemode='a', level=logging.DEBUG)

print "'" + sys.path[0] + "'"  # This will give the parent folder path of the current executing file. currently not used within the script.
SCRIPT_PATH = sys.path[0]

# Input Workbook Name - DATA_INPUT_WORKBOOK_NAME variable present in Config_ExpenCalc.py 
INPUT_FILE=DATA_INPUT_WORKBOOK_NAME 

# Report File Name - EXPENCALC_REPORT variable present in Config_ExpenCalc.py
#REPORT_FILE=EXPENCALC_REPORT

logging.info('Starting data crunching program...')

if not (os.path.isfile(SCRIPT_PATH + INPUT_FILE)):
	print "Input file - " + INPUT_FILE + ' not found.'
	logging.critical("Input file - " + INPUT_FILE + ' not found.')
	quit()

year = str(datetime.datetime.now().year) # year is returned in int hence typecast to string
D_Grph_Yearly_data = {}

col_width = 256 * 20             # 20 characters wide

ezxf = xlwt.easyxf
heading_xf = ezxf('font: bold on; align: wrap on, vert centre, horiz center; pattern: pattern solid, fore-colour grey25')
##color_xf = ezxf('pattern: pattern solid, fore-colour ice_blue')

style = xlwt.XFStyle()
style.num_format_str = '#,##0.00'

workbook = xlrd.open_workbook(SCRIPT_PATH + INPUT_FILE)
wb = xlwt.Workbook()
print "Number of sheets - " , workbook.nsheets

# --- Logic to skip calculation on worksheet which has 'NO' keyword at the end and crunch stats for each sheet. -----
for vInvSheet in workbook.sheets():
    if vInvSheet.name.find("NO") != -1:
        print "Skipping sheet - " , vInvSheet.name
        logging.info("Skipping sheet - %s" , vInvSheet.name)
    else:  
	print "Sheet  - " , vInvSheet.name       
	logging.info("Sheet  - %s" , vInvSheet.name) 
	Testworksheet = workbook.sheet_by_name(vInvSheet.name)

	#Testworksheet = workbook.sheet_by_name(year)  # datetime.datetime.now().year   --- returns the current year

# ----- Start of Sorting entire excel based on Date column -----
	logging.info('Start of Sorting...')
    # Sorting of any column, just give the column number to target_column variable
	target_column = 0 # Sort on Date field

	data = [Testworksheet.row_values(i) for i in xrange(Testworksheet.nrows)] # returns list with rows with a list of columns

	labels = data[0]  # Header row

	data = data[1:]  # Complete data except the header

# sort the data based on Date field
    ##data.sort(key=lambda x: x[target_column])

	data.sort(key = itemgetter(target_column))  # Complete sorted data in the variable 'data' based on the 'Date' column
	# >>> itemgetter(1)('ABCDEFG')
	# 'B'
	# >>> itemgetter(1,3,5)('ABCDEFG')
	# ('B', 'D', 'F')

	logging.info('End of Sorting...')
# ----- End of Sorting -----

# ----- Calculation of Income -----
	L_Income = []
	D_Income = {}

	D_SumOfUniqueTags = {}
	D_Deduct = {}

	logging.info('Starting to read each row and sum up values for each Tag even if its duplicate....')

	for tag in data:
	    #print tag
	    date_tuple = xlrd.xldate_as_tuple(tag[0],workbook.datemode)
	    now = datetime.datetime(date_tuple[0] , date_tuple[1], date_tuple[2])

	    if tag[3] == 'Income':   # Calculation of Income, appending to the Income list
	        L_Income.append((now.month, tag[2]))  # tag[2] is the 'Price' variable, L_Income structure is [(Month, Value), (1, 2353.62), (1, 42), (2, 3042.11)]
	        # logging.info(L_Income)

	    else:    # All the other tags except Income
	        L_Rowtags = tag[3].split(',')    # split if you have multiple tags in one row. Split returns a list of items
		L_Rowtags = [catg.strip() for catg in L_Rowtags] # Trimming spaces for each of the multiple tag after split done above
   
		DeductAmt = 0	# Clears previous stored amount
		# below loop for Total of each unique Tag for entire input not based on month or date
		for newTag in L_Rowtags:   
            
		    RetVal = newTag.find("#") 
		    # Below If block along with DeductAmt and D_Deduct is part of 1.1 release
		    if RetVal != -1:   # -1 means substring not found; if you get # in the tagname, then strip the # from the end of the tagname
		        newTag = newTag[0:len(newTag)-1]
			#print newTag
			
			DeductAmt = float(DeductAmt) + float(tag[2])	 # tag[2] is the price 	
			# D_Deduct dict variable holds the total of all expense tags for each month. This will hold those prices whose tag has # so that it is removed from calculation 	
			if D_Deduct.has_key(now.month):
			    D_Deduct[now.month] = float(D_Deduct[now.month]) + float(tag[2])
			else:
		            D_Deduct[now.month] = float(tag[2])
            # Total of each unique Tag for entire input not based on month or date
		    if D_SumOfUniqueTags.has_key((now.month, newTag)):     # if tags are present in dictionary sum the price
			#print newTag , ' -- ', tag[2]
			#print D_SumOfUniqueTags[(now.month, newTag)]
		        D_SumOfUniqueTags[(now.month, newTag)] = float(D_SumOfUniqueTags[(now.month, newTag)]) + float(tag[2])

		    else:   # if new tags then add new tag to dictionary and

                	D_SumOfUniqueTags[(now.month, newTag)] = float(tag[2])

##print D_SumOfUniqueTags.items()

# ------- Group the sorted output by date ---------
	# The itertools.groupby() function takes a sequence and a key function, and returns an iterator that
    # generates pairs. Each pair contains the result of key_function(each item) and another iterator containing
    # all the items that shared that key result.

    # L_Income structure is ('Month','Price'), group this below
	# logging.info('Pls show me -- %s', str(D_Deduct))
	#print D_Deduct
	groups = groupby(L_Income, itemgetter(0)) # Hence, itemgetter[0] here means Month, 1st column
# This was grouping of all Income based data based on date column.

# ------- End of Grouping ---------
# ---- Iterate the Income Group and get the sum of all individual incomes for a single month. -----
	for key,value in groups:
	    s = sum([ float(item[1]) for item in value ])
	    D_Income[key] = s  # store the monthly(variable key) summed income(variable s) into dictionary to match later
	    # logging.info('Ikey %s -- Ival %s', str(key), str(s))
	    print "Income for month " , key , " is amount " , s
	logging.info('End of Calculation of Income...')
# ----- End of Calculate Income -----


# ------- Group the sorted excel output minus the header done above in 'data' variable by date ---------

	groups_Data = groupby(data, itemgetter(target_column))

# ------- End of Grouping ---------

# ------- Calculate Date wise Total --------
	T_DailyTotal = () # Tuple to store pair of date and total expense
	L_DailyTotal = []

	for k1,v1 in groups_Data:
	    s = sum([ float(item[2]) for item in v1 ]) # item[2] is Rate field
	    date_tuple = xlrd.xldate_as_tuple(item[0],workbook.datemode)
	    now = datetime.datetime(date_tuple[0] , date_tuple[1], date_tuple[2])
    
	    T_DailyTotal = item[0] , round(s,2) , now.month # item[0] is the date field, s is the daily total, month for this date
	    #print T_DailyTotal
	    L_DailyTotal.append(T_DailyTotal)

# ------- End of Calculate Date wise Total --------

# ------- Calculate Month wise Total --------

		#ws = wb.add_sheet('SummaryExpense')
	ws = wb.add_sheet(vInvSheet.name)
	month_headings = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC','Total','%']
	rowx = 1
	for colx, value in enumerate(month_headings):
            ws.write(rowx, colx+1, value, heading_xf)

	ws.col(0).width = col_width # Column Width
	ws.write(2,0,'INCOME')
	ws.write(3,0,'TOTAL EXPN')
	ws.write(4,0,'SAVINGS')

	IncomeTotal = 0
	TotalExpense = 0
	TotalSavings = 0
	
	groups_DailyTot = groupby(L_DailyTotal, itemgetter(2)) # Group the L_DailyTotal list by month

	# Fixed the Income bug, Bug1self, by checking, if income is available for the month, if not, return 0
	# logging.info(D_Deduct)
	# Calculating below monthly Expense, Income and Savings
	# D_Deduct dict variable holds the total of all expense tags for each month 
	for k,v in groups_DailyTot:
	    
	    s = sum([ item[1] for item in v ]) # item[1] is the daily total. Hence, variable s is the total of all daily total
	    varMonth = k
	    # logging.info('Emonth %s -- Eval %s', str(varMonth), str(s))
	    # below If block is part of 1.1 release, this is to remove the price which has to be excluded based on # appended to tag
	    if D_Deduct.has_key(k):
	        varExpense = (s-D_Income.get(k, 0)) - D_Deduct.get(k)
	        # logging.info('check-----')
	    else:
		varExpense = (s-D_Income.get(k, 0))
	    
            #print varExpense
	    ws.write(2,k,D_Income.get(k, 0), style)  # Income
	    IncomeTotal = IncomeTotal + D_Income.get(k, 0)

	    ws.write(3,k,varExpense, style)     # TotalExpense
	    TotalExpense = TotalExpense + varExpense

	    saved = (D_Income.get(k, 0)- varExpense)

	    ws.write(4,k,saved, style)          # Savings
	    TotalSavings = TotalSavings + saved

	    print varMonth , "th month Income =" , D_Income.get(k, 0), "Expense =" , varExpense, "Savings =" ,saved # income to be minus from expense
	    # logging.info('click-----')

	ws.write(2,13,IncomeTotal,style)   # Income Total
	ws.write(3,13,TotalExpense,style)  # Expense Total
	ws.write(4,13,TotalSavings,style)  # Savings Total

	# --- Below if else block only to capture data for graph ----
	# if len(D_Grph_Yearly_data)!= 0: # If not empty ,then add to the dict  
	D_Grph_Yearly_data.update({str(vInvSheet.name):(IncomeTotal,TotalExpense,TotalSavings)}) 
	# D_Grph_Yearly_data = D_Grph_Yearly_data + {str(vInvSheet.name):(IncomeTotal,TotalExpense,TotalSavings)}
	# else:	
		# If no data exist then create one entry
		# D_Grph_Yearly_data = {str(vInvSheet.name):(IncomeTotal,TotalExpense,TotalSavings)}

	print D_Grph_Yearly_data

# ------- End of Calculate Month wise Total --------    

	DupTag = [tags for mnth,tags in D_SumOfUniqueTags.keys()]
	L_UniqueTag = list(set(DupTag)) # removes duplicate without considering Order.

	row = 6
	col = 0
	percentTotal = 0

	for Category in L_UniqueTag:  # Print category list into Excel
	    ws.write(row,col,Category)
	    catgTotal = 0
    
	    for mnth,tags in D_SumOfUniqueTags.keys():      # For each month, it will insert values for the selected category
	        if Category == tags:
		    ws.write(row,col+mnth,D_SumOfUniqueTags[(mnth, tags)], style) #Print values against category in appropriate month
		    catgTotal = catgTotal + D_SumOfUniqueTags[(mnth, tags)]
		            
	    percentTotal = ( catgTotal / IncomeTotal ) * 100      # Calculate percent against total income for each category
	    ws.write(row,13,catgTotal,style)
	    ws.write(row,14,percentTotal,style)                     
	    row = row+2

	ws.panes_frozen = True
	ws.horz_split_pos = 5
	ws.vert_split_pos = 1

wb.save(SCRIPT_PATH + '/ExpenseSummary' + year + '.xls')
logging.info('END OF EXPN-CALC.')

# --- HTML Reporting, expecting javascript file to be in the cwd ---
grph_Col_year_count = len(D_Grph_Yearly_data)
year = str(year)
if grph_Col_year_count == 0:
	html_str = """<!DOCTYPE HTML>
		<html>
		<head></head>
		<body>
		<h1>There is no Yearly data available..."</h1>
		</body></html>"""
else: 
	pie_title_currYear = ' title:{text: "' + year + ' Income - ' + str(D_Grph_Yearly_data[year][0]) + '"},'  # Requirement to only see pie for current year, variable year always holds current year.
	pie_datapoints = '{ y: ' + str(D_Grph_Yearly_data[year][1]) + ', name: "Expense", legendMarkerType: "square"}, { y: ' + str(D_Grph_Yearly_data[year][2]) + ', name: "Savings", legendMarkerType: "circle"} '
	num_year = int(year)
	Expense_label = ''
	Savings_label = ''
	Income_label = ''
	for item_num in range(grph_Col_year_count):  
		print item_num
		if item_num < 2: # Restricting graphing data to previous 2 years only.
			#if D_Grph_Yearly_data[str(num_year)][2] < 0
			Expense_label = Expense_label + '{label: "' + str(num_year) + '", y: ' + str(D_Grph_Yearly_data[str(num_year)][1]) + '},'
			Savings_label = Savings_label + '{label: "' + str(num_year) + '", y: ' + str(D_Grph_Yearly_data[str(num_year)][2]) + '},'
			Income_label = Income_label + '{label: "' + str(num_year) + '", y: ' + str(D_Grph_Yearly_data[str(num_year)][0]) + '},'
			num_year = num_year - 1  # Logic to select only the current year and previous 2 year data
			logging.info("Collating graphing data...")
		else:
			break
	
	Expense_label = Expense_label[0:len(Expense_label)-1]   # To remove the trailing comma
	Savings_label = Savings_label[0:len(Savings_label)-1]   # To remove the trailing comma
	Income_label = Income_label[0:len(Income_label)-1]   # To remove the trailing comma
			
	html_str = """<!DOCTYPE HTML>
		<html>
		<head>
		<script type="text/javascript">
		window.onload = function () {
			var chart = new CanvasJS.Chart("chartContainer",
			{
				""" + pie_title_currYear + """								
				data: [
				{
					type: "pie",
					indexLabelFontFamily: "Garamond",
					indexLabelFontSize: 20,
					indexLabelFontWeight: "bold",
					startAngle:0,
					indexLabelFontColor: "MistyRose",
					indexLabelLineColor: "darkgrey",
					indexLabelPlacement: "inside",
					toolTipContent: "{name}: {y}",
					showInLegend: true,
					indexLabel: "#percent%",
					dataPoints: [
						""" + pie_datapoints + """
					]
				}
			      ]
			});
			chart.render();
			
			var chart = new CanvasJS.Chart("chartContainer2",
			{
			theme: "theme2",
                        animationEnabled: true,
			title:{
				text: "Income Vs Expense Vs Savings Trend",
				fontSize: 30
			},
			toolTip: {
				shared: true
			},
			axisX:{
				title: "Interactive Report" ,
				tickColor: "red",
				tickLength: 15,
				tickThickness: 3
			},

			axisY: {
				title: "GBP" ,
				 tickLength: 15,
				 tickColor: "DarkSlateBlue" ,
				 tickThickness: 5
				//valueFormatString:  "##,#0", // move comma to change formatting
				//prefix: "$"
			},
			//axisY2: {
				//title: "Great Britian Pounds"
			//},

			legend:{
				verticalAlign: "top",
				horizontalAlign: "center"
			},
			data: [ 
			{
				type: "column",	
				name: "Expense",
				legendText: "Expense",
				showInLegend: true, 
				dataPoints:[
				""" + Expense_label + """
				]
			},
			{
				type: "column",	
				name: "Income",
				legendText: "Income",
				showInLegend: true,
				dataPoints:[
				""" + Income_label + """
				]
			},
			{
				type: "column",	
				name: "Savings",
				legendText: "Savings",
				//axisYType: "secondary",
				showInLegend: true,
				dataPoints:[
				""" + Savings_label + """
				]
			}
			
			],
          legend:{
            cursor:"pointer",
            itemclick: function(e){
              if (typeof(e.dataSeries.visible) === "undefined" || e.dataSeries.visible) {
              	e.dataSeries.visible = false;
              }
              else {
                e.dataSeries.visible = true;
              }
            	chart.render();
            }
          },
        });

		chart.render();

		}
		</script>
		<script type="text/javascript" src="canvasjs.min.js"></script>
		</head>
		<body>
			<div id="chartContainer" style="height: 300px; width: 50%;"></div>
			<div>
		    <div id="chartContainer2" style="height: 400px; "></div>
		    </div>
		</div>	
		</body>
		</html>
		"""	

Html_file = open(SCRIPT_PATH + '/Report.html',"w")
Html_file.write(html_str)
Html_file.close()
webbrowser.open_new_tab(SCRIPT_PATH + '/Report.html')
