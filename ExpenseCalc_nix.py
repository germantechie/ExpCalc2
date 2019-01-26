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
#       and multi-column bar chart to show trend.                                           |
# 2.0 - 25 October 2015                                                                     |
#       Rewritten the entire codebase to more concise and less code retaining all the       |
#       functionality except for rev 1.1 w.r.t #. Not included in this revision.            |
# All further versions will be through Git.                                                 |
# -------------------------------------------------------------------------------------------

import xlrd, xlwt
import datetime, sys, os
import logging, subprocess, webbrowser

from itertools import groupby
from operator import itemgetter

# Config_ExpenCalc.py is a configuration file which should be in the current directory as this script
from Config_ExpCalc import * 

class varSettings(object):
	SCRIPT_PATH = sys.path[0]  # This will give the parent folder path of the current executing file. currently not used within the script.

	# Input Workbook Name - DATA_INPUT_WORKBOOK_NAME variable present in Config_ExpenCalc.py 
	INPUT_FILE=DATA_INPUT_WORKBOOK_NAME   #'Exp_Input.xls'     

	# Report File Name - EXPENCALC_REPORT variable present in Config_ExpenCalc.py
	#REPORT_FILE=EXPENCALC_REPORT

	year = str(datetime.datetime.now().year) # year is returned in int hence typecast to string
	D_Grph_Yearly_data = {}

	col_width = 256 * 20             # 20 characters wide

	ezxf = xlwt.easyxf
	heading_xf = ezxf('font: bold on; align: wrap on, vert centre, horiz center; pattern: pattern solid, fore-colour grey25')
	##color_xf = ezxf('pattern: pattern solid, fore-colour ice_blue')

	style = xlwt.XFStyle()
	style.num_format_str = '#,##0.00'

	D_htmlTableData = {}

	L_userReportTags = ['Food', 'Travel', 'Bills']  # offering user defined TAG based report
	D_grpTagReport = {}  # Empty Dictionary

	def Read_Input(self):
		self.Inp_filepath = os.path.join(VARs.SCRIPT_PATH, VARs.INPUT_FILE)
		#print(self.Inp_filepath)
		if not (os.path.isfile(self.Inp_filepath)):
			print("Input file - " + self.Inp_filepath + ' not found. Please check your Config file...')
			#logging.critical("Input file - " + INPUT_FILE + ' not found.')
			sys.exit(0)

		self.workbook = xlrd.open_workbook(self.Inp_filepath)   # To Read Excel
		self.wb = xlwt.Workbook()								# To Write Excel
	#print "Number of sheets - " , workbook.nsheets

def group_the_Tags_monthly(TAG, PRICE, MONTH):
	''' This method groups the tag and totals the amount for that tag while noting the month.
	    Going with this below code such that report based tags can be fetched from single dictionary
		'''
	if (TAG, MONTH) in VARs.D_grpTagReport:
		VARs.D_grpTagReport[TAG, MONTH] += PRICE
	else:
		VARs.D_grpTagReport[TAG, MONTH] = PRICE

def Crunch_Data(MSTR_DATA):
	''' 
	This function is used to crunch all the input data and store into respective dictionary variables to pass
	those to the reporting functions.
	'''
	deduct_AMT = 0
	D_Total_Monthly_Expense = {}
	D_monthly_dedcut = {1:0,2:0,3:0,4:0,5:0,6:0,7:0,8:0,9:0,10:0,11:0,12:0}

	for tag in MSTR_DATA:  # Loop on the sorted data row by row.
		date_tuple = xlrd.xldate_as_tuple(tag[0],VARs.workbook.datemode)
		rwDate = datetime.datetime(date_tuple[0] , date_tuple[1], date_tuple[2])

		dMonth = rwDate.month
		TAG = tag[3]
		PRICE = float(tag[2])

		if TAG.find(",") == -1:    # comma not found
			group_the_Tags_monthly(TAG, PRICE, dMonth)

		else:
			L_Rowtags = TAG.split(',')  # get seperate tags
			L_Rowtags = [catg.strip() for catg in L_Rowtags] # Trimming spaces for each of the multiple tag after split done above
			#pdb.set_trace()
			iter_Count = len(L_Rowtags) - 1
			dupl_amt = (iter_Count * PRICE)
			deduct_AMT = dupl_amt + deduct_AMT		# Keep a tab of total deduction bcz by splitting each of the tags have the same amount which is duplicate in calculation
			
			D_monthly_dedcut[dMonth] += dupl_amt 	# Keep a tab of monthly deduction 
			
			for tg in L_Rowtags:
				group_the_Tags_monthly(tg, PRICE, dMonth)

	D_Monthly_Tag_Values = VARs.D_grpTagReport.copy()   # giving a logical name  {(Tag1, month):amt, (Tag2, month):amt}
	#pdb.set_trace()
	# Calculate Monthly Expense Total 
	D_Total_Monthly_Expense = { n[1]:0 for n in VARs.D_grpTagReport.keys() }  # writing {month:0} with n[0]	

	# Calculate Total for the Tags based on monthly aggregate -- http://stackoverflow.com/questions/33252985/how-to-sum-values-from-a-python-dictionary-having-key-as-tuple
	# writing {Tag:0} with k[0]
	D_Total_Tag_Values = { k[0]:0 for k in VARs.D_grpTagReport.keys() }  # Get all keys by updating its value to 0 so it overwrites and gives unique key.

	for key in VARs.D_grpTagReport:	# Loop through the monthly values and add the value with the new dict value
		D_Total_Tag_Values[key[0]] = D_Total_Tag_Values[key[0]] + VARs.D_grpTagReport[key]
		D_Total_Monthly_Expense[key[1]] = D_Total_Monthly_Expense[key[1]] + VARs.D_grpTagReport[key]
		#print(key, D_Total_Monthly_Expense[key[1]])
	
	#print(D_Total_Tag_Values)  # result is {'Bills': 1577.0399999999997, 'Food': 92.48}
	# print(D_Total_Monthly_Expense)  # result is {'1': 157734, '2': 923453.48}
	return (D_Monthly_Tag_Values, D_Total_Tag_Values, D_Total_Monthly_Expense, deduct_AMT, D_monthly_dedcut)

def write_EXCEL_Report(vInvSheet, D_Monthly_Tag_Values, D_Total_Tag_Values, D_Total_Monthly_Expense, deduct_AMT, D_monthly_dedcut):
	'''
	This function accepts the dictionary variables and reads the required data to present it in Excel.
	'''
	style = VARs.style
	ws = VARs.wb.add_sheet(vInvSheet.name)
	month_headings = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC','Total','% of Inc']
	rowx = 1
	for colx, value in enumerate(month_headings):
            ws.write(rowx, colx+1, value, VARs.heading_xf)  # writing the headers to excel

	ws.col(0).width = VARs.col_width # Column Width
	# writing the headers to excel
	ws.write(2,0,'INCOME')
	ws.write(3,0,'EXPENSE')
	ws.write(4,0,'SAVINGS')
	
	if 'income' in D_Total_Tag_Values: 
		IncomeTotal = D_Total_Tag_Values['income']  # check if income is the only word, data cleansing required
	else:
		IncomeTotal = 0 
	TotalExpense = sum(D_Total_Monthly_Expense.values()) - IncomeTotal - deduct_AMT   # This will sum all the values within the dict as it comprises {month:amount}
	TotalSavings = IncomeTotal - TotalExpense
	
	if IncomeTotal == 0:
		SavingsPercent = 0.00
	else:
		SavingsPercent = (TotalSavings/IncomeTotal)*100
    
	# writing the above totals to excel
	ws.write(2,13,IncomeTotal,style)   # Income Total
	ws.write(3,13,TotalExpense,style)  # Expense Total
	ws.write(4,13,TotalSavings,style)  # Savings Total	
	ws.write(4,14,SavingsPercent,style)  # Savings Percent of Income
 
	# Print category list into Excel >>>>>>>>>
	row = 4 
	col = 0
	percentTotal = 0
	D_catg = {}

	for mn in D_Total_Monthly_Expense:
		# check if that month has income, if yes then deduct that amount from expense bcz income was also one of the tag for calculation
		MINUS = D_monthly_dedcut[mn]
		monthly_expense = D_Total_Monthly_Expense[mn] - MINUS

		if ('income', mn) in D_Monthly_Tag_Values:
			monthly_income = D_Monthly_Tag_Values['income',mn]
			monthly_expense = monthly_expense - monthly_income  # Expense Tag Total minus any income for that month
			monthly_savings = monthly_income - monthly_expense  

			ws.write(2, col+mn, monthly_income, style)			 		 # Enter Income Tag for that month
			ws.write(3, col+mn, monthly_expense, style)	 # Enter Expense Tag Total for that month minus any income
			ws.write(4, col+mn, monthly_savings, style)			 		 # Enter Savings Tag for that month
			
			monthly_income = monthly_expense = monthly_savings = 0       # Clear Variable values
		else:
			ws.write(2, col+mn, 0, style)								 # Enter Income Tag as 0 for that month
			ws.write(3, col+mn, monthly_expense, style)					 # Enter Expense Tag Total for that month
			ws.write(4, col+mn, 0-monthly_expense, style)			 	 # Enter Savings Tag Total for that month
			monthly_expense = 0

	HTML_rptTable = '<tr>'
	htmlRpt_MNTH = 0
	HTML_Total = ''
	HTML_Percent = ''
	for catg, mnth in sorted(D_Monthly_Tag_Values):  # this will sort the dictionary based on key -- mnth is an integer
		if catg != 'income':		# Don't print Income TAG
			if catg not in D_catg:  # just a means to distinguish new Tags
				row += 2 			# Leaving one row blank as styling
				D_catg[catg] = 0 
				
				if htmlRpt_MNTH != 0:   # to avoid creating empty row.
					#print(vInvSheet.name, catg, htmlRpt_MNTH) 
					while (htmlRpt_MNTH != 12):   # While loop will print blank td's for blank months.
						HTML_rptTable = HTML_rptTable + '<td></td>'
						htmlRpt_MNTH = htmlRpt_MNTH + 1
					HTML_rptTable = HTML_rptTable + '<td>' + HTML_Total + '</td><td>' + HTML_Percent + '</td>' + '</tr><tr>'

				HTML_rptTable = HTML_rptTable + '<td>' + catg.title() + '</td>'  # First Column should be Tag Name

				# ---- MSG1 -- If new category does not have entry from Jan, then enter empty columns until this category has an entry.
				cntMnth = mnth          				# mnth always starts from 1
				while (cntMnth >= 2):   				# While loop will print blank td's for initial blank months.
					HTML_rptTable = HTML_rptTable + '<td></td>'
					cntMnth = cntMnth - 1
				# ----- End MSG1 -------------------------------------------------------

				# ALL the below code runs once for a unique TAG
				if IncomeTotal == 0:       # Cannot calculate percent of Income as it will be division by zero
					percentTotal = "NA" 
				else:
					percentTotal = ( D_Total_Tag_Values[catg] / IncomeTotal ) * 100      # Calculate percent against total income for each category
				ws.write(row, 13, D_Total_Tag_Values[catg], style)					 # Enter Tag Total for whole year
				ws.write(row, 14, percentTotal, style)								 # Enter Percent of Income
				ws.write(row, col, catg.title(), style)    							 # Enter Tags

			ws.write(row, col+mnth, D_Monthly_Tag_Values[catg, mnth], style)	        # Enter Tag Value for that month
			
			# ---- MSG2 ------ To enter blank columns, if category does not have values for in between months. 
			if htmlRpt_MNTH != 0:
				mnth_diff = mnth - htmlRpt_MNTH
				if mnth_diff > 1:
					#fill empty td tags
					for u in range(mnth_diff-1): 
						HTML_rptTable = HTML_rptTable + '<td></td>' 
			htmlRpt_MNTH = mnth  # To save the last month in the loop, based on this info, need to insert empty TD tags.
			# ----End of MSG2 ------------------------------------------------------ 
			
			HTML_rptTable = HTML_rptTable + '<td>' + str(D_Monthly_Tag_Values[catg, mnth]) + '</td>'   # Enter the month's value for the category
			
			HTML_Total = str(D_Total_Tag_Values[catg])
			HTML_Total = str(format(float(HTML_Total), ',.2f'))
			if percentTotal == "NA":
				HTML_Percent = "NA"
			else:
				HTML_Percent = str(round(percentTotal,2))

	while (htmlRpt_MNTH != 12):   # While loop will print blank td's for blank months for last row.
		HTML_rptTable = HTML_rptTable + '<td></td>'
		htmlRpt_MNTH = htmlRpt_MNTH + 1

	HTML_rptTable = HTML_rptTable + '<td>' + HTML_Total + '</td><td>' + HTML_Percent + '</td>' + '\n'
	HTML_rptTable = HTML_rptTable + '</tr>'	
	VARs.D_htmlTableData[vInvSheet.name] = HTML_rptTable   # Dicton{'2015_uk': 'html_data'} storing tabular data of each year in Global Class variable with Year (sheet name) as the key

	ws.panes_frozen = True
	ws.horz_split_pos = 5
	ws.vert_split_pos = 1

def write_HTML_Report():
	'''
	This function should help generate static html report.
	'''
	rptHEAD = ''' 
	<!DOCTYPE>
	<html>
	  <head>
	    <meta charset="utf-8">
	    <title>Net Worth Visualizer</title>
	    <meta name="viewport" content="width=device-width, initial-scale=1">
	    
	    <link rel="stylesheet" type="text/css" href="dist/css/bootstrap.css">
	'''
	rptHEAD_JS_static_start = '''
	<script type="text/javascript">
	      window.onload = function () {

	        var chart = new CanvasJS.Chart("chartContainer1",
	        {

	          title:{
	            text: "Expense Across Categories",
	            fontFamily: "sans-serif",
	            fontSize: 22
	          },
	                            animationEnabled: true,
	          axisX:{

	            gridColor: "Silver",
	            tickColor: "silver",
	            valueFormatString: "MMM/YY"

	          },                        
	          toolTip:{  shared:true
	                            },
	          theme: "theme3",
	          axisY: {
	            gridColor: "Silver",
	            tickColor: "silver"
	          },
	          legend:{
	            verticalAlign: "center",
	            horizontalAlign: "right"
	          },
	          data: [
	          {        
	            type: "line",
	            showInLegend: true,
	            lineThickness: 2,
	            name: "Food",
	            markerType: "square",
	            color: "#F08080",
	            dataPoints: [
	            { x: new Date(2014,7,01), y: 650 },
	            { x: new Date(2014,8,01), y: 700 },
	            { x: new Date(2014,9,01), y: 710 },
	            { x: new Date(2014,10,01), y: 658 },
	            { x: new Date(2014,11,01), y: 734 },
	            { x: new Date(2015,0,01), y: 963 },
	            { x: new Date(2015,1,01), y: 847 },
	            { x: new Date(2015,2,01), y: 853 },
	            { x: new Date(2015,3,01), y: 869 },
	            { x: new Date(2015,4,01), y: 943 },
	            { x: new Date(2015,5,01), y: 970 }
	            ]
	          },
	          {        
	            type: "line",
	            showInLegend: true,
	            name: "Travel",
	            color: "#20B2AA",
	            lineThickness: 2,

	            dataPoints: [
	            { x: new Date(2014,7,01), y: 510 },
	            { x: new Date(2014,8,01), y: 560 },
	            { x: new Date(2014,9,01), y: 540 },
	            { x: new Date(2014,10,01), y: 558 },
	            { x: new Date(2014,11,01), y: 544 },
	            { x: new Date(2015,0,01), y: 693 },
	            { x: new Date(2015,1,01), y: 657 },
	            { x: new Date(2015,2,01), y: 663 },
	            { x: new Date(2015,3,01), y: 639 },
	            { x: new Date(2015,4,01), y: 673 },
	            { x: new Date(2015,5,01), y: 660 }
	            ]
	          },
	          {        
	            type: "line",
	            showInLegend: true,
	            name: "Bills",
	            color: "#26FG3A",
	            lineThickness: 2,

	            dataPoints: [
	            { x: new Date(2014,7,01), y: 310 },
	            { x: new Date(2014,8,01), y: 460 },
	            { x: new Date(2014,9,01), y: 240 },
	            { x: new Date(2014,10,01), y: 758 },
	            { x: new Date(2014,11,01), y: 544 },
	            { x: new Date(2015,0,01), y: 593 },
	            { x: new Date(2015,1,01), y: 357 },
	            { x: new Date(2015,2,01), y: 963 },
	            { x: new Date(2015,3,01), y: 439 },
	            { x: new Date(2015,4,01), y: 373 },
	            { x: new Date(2015,5,01), y: 760 }
	            ]
	          }

	          
	          ],
	              legend:{
	                cursor:"pointer",
	                itemclick:function(e){
	                  if (typeof(e.dataSeries.visible) === "undefined" || e.dataSeries.visible) {
	                    e.dataSeries.visible = false;
	                  }
	                  else{
	                    e.dataSeries.visible = true;
	                  }
	                  chart.render();
	                }
	              }
	        });

	    chart.render();

	      var chart = new CanvasJS.Chart("chart_NetWorth",
	      {
	        title:{
	          text: null,
	        },     
	            animationEnabled: true,     
	                    
	        data: [
	        {        
	          type: "doughnut",
	          startAngle: 100,                          
	          toolTipContent: "{legendText}: {y} - <strong>#percent% </strong>",          
	          showInLegend: false,
	          dataPoints: [
	            {y: 65899660, indexLabel: "Total Assets #percent%", legendText: "Total Assets" },
	            {y: 60929152, indexLabel: "Total Liability #percent%", legendText: "Total Liability" } 
	          ]
	        }
	        ]
	      });
	      chart.render();
	      document.getElementById("ch_dough_center").innerHTML = "Net Worth 5.5 Lac";

	      }
	    </script>

	</head>
	'''

	rptBody_1 = '''
	<body>
    <div class="jumbotron_Header"> <!-- Top Header -->
      <h3 style="padding-left:20px; color:white;">Net Worth Visualizer</h3>
      <section style="padding-right:20px;" class="pull-right">
          <span class="label label-primary">Income <big><span style="font-size:medium">&#8377; 12,3953</big></span>
          <span class="label label-success">Savings <span style="font-size:medium">&#8377; 12,3853</span></span>
          <span class="label label-warning">Expense <span style="font-size:medium">&#8377; 6,874.56</span></span>
      </section>
    </div> 
    
    <div class="container-fluid"> <!-- Rest of the Body -->
      <nav>  <!-- Navigation -->
        <ul class="nav nav-pills">
        	<li role="presentation" class="active"><a data-toggle="pill" href="#home">Home</a></li>
        '''
	
	for ids in VARs.D_htmlTableData.keys():
 		rptBody_1 = rptBody_1 + '''
    	<li role="presentation"><a data-toggle="pill" href="#menu_''' + ids + '''">''' + ids + '''</a></li>''' 
          
	rptBody_1 = rptBody_1 + '''      
        </ul>
      </nav>    

      <div class="tab-content">   <!-- Tab Content tag should be only ONE.-->
          <div id="home" class="tab-pane fade in active">     <!-- Home Contents Area -->
            <p></p> 
               <div class="row row-eq-height">
                  <div class="col-xs-7">
                     <p>
                      <div id="chart_NetWorth" style="height: 300px; width: 100%;"></div>
                      <div id="ch_dough_center" style="position:absolute;left:1px;top:15px;height:50%;width:100%;line-height:260px;text-align:center;color:black;font-size:16px;white-space: pre-wrap;">300</div>
                    </p>
                  </div>
                  <!-- <div style="width:3px; height:250px; margin:20px 40px 40px 0px; background-color:#808187;"></div> -->
                  <div class="col-xs-5"> <!-- style="background-color: #dedef8;box-shadow: inset 1px -1px 1px #444, inset -1px 1px 1px #144;" -->
                      <div class="col-xs-6">  
                        <h4 class="sub-header"><strong>Total Assets</strong></h4>
                        <div class="table-responsive">
                          <table class="table table-striped" style="font-family:Helvetica Neue, Helvetica, Arial, sans-serif; font-size: 12px;">
                            <tbody>
                              <tr>
                                <td>Bank Savings</td>
                                <td>Rs. 23,567</td>
                              </tr>
                               <tr>
                                <td>Bank FD's</td>
                                <td>Rs. 23,567</td>
                              </tr>
                              <tr>
                                <td>Mutual Funds</td>
                                <td>Rs. 23,567</td>
                              </tr>
                              <tr>
                                <td>Gold</td>
                                <td>Rs. 23,567</td>
                              </tr>
                              <tr>
                                <td>CAR</td>
                                <td>Rs. 23,567</td>
                              </tr>
                              <tr>
                                <td>Personal Goods</td>
                                <td>Rs. 23,567</td>
                              </tr>
                              <tr>
                                <td>LIC</td>
                                <td>Rs. 23,567</td>
                              </tr>
                            </tbody>
                          </table>
                        </div>
                      </div>
                      <div class="col-xs-6">
                        <h4 class="sub-header"><strong>Total Liability</strong></h4>
                        <div class="table-responsive">
                          <table class="table table-striped" style="font-family:Helvetica Neue, Helvetica, Arial, sans-serif; font-size: 12px;">
                            <tbody>
                              <tr>
                                <td>Home Loan</td>
                                <td>Rs. 23,567</td>
                              </tr>
                              <tr>
                                <td>Flat Expense</td>
                                <td>Rs. 23,567</td>
                              </tr>
                            </tbody>
                          </table>
                        </div>
                      </div>
                  </div>  <!-- End of Line 134 and second half of home page-->
               </div>
            
              <div class="row row-eq-height">
                  <div class="col-xs-2">
                    <p></p>
                  </div>
                  <div class="col-xs-8">
                      <p>
                          <div id="chartContainer1" style="height: 300px; width: 100%;"></div>
                      </p>
                  </div>
                  <div class="col-xs-2">
                    <p></p>
                  </div>
              </div> 

          </div> <!-- End of Home Page Tab-->
          '''
 	for ids in VARs.D_htmlTableData.keys():
		rptBody_1 = rptBody_1 + '''<div id="menu_''' + ids + '''" class="tab-pane fade"> <!-- new Tab Area -->  '''
		rptBody_1 = rptBody_1 + '''  
            <h3>Menu 1</h3>
            <p>Some content in menu 1.</p>
            <table class="table table-striped">
              <thead>
                  <tr>
                    <th>Category</th>
                    <th>Jan</th>
                    <th>Feb</th>
                    <th>Mar</th>
                    <th>Apr</th>
                    <th>May</th>
                    <th>Jun</th>
                    <th>Jul</th>
                    <th>Aug</th>
                    <th>Sep</th>
                    <th>Oct</th>
                    <th>Nov</th>
                    <th>Dec</th>
                    <th>Total</th>
                    <th>% Income</th>
                  </tr>
              </thead>
              <tbody>
			'''
		rptBody_1 =  rptBody_1 + VARs.D_htmlTableData[ids]
		rptBody_1 =  rptBody_1 + '''</tbody>
            </table>
          </div>
          '''


	rptTest_Last = '''

      </div>
    </div>
	   <script src="dist/js/jquery.min.js"></script>
     <script src="dist/js/bootstrap.min.js"></script>
     <script src="dist/js/canvasjs.min.js"></script>

  </body>
</html>
	'''
	fo = open(os.path.join(VARs.SCRIPT_PATH, "HTML_Report.html"), "w")
	fo.write(rptHEAD + rptHEAD_JS_static_start + rptBody_1 + rptTest_Last)
	fo.close()

# ----------------------------  MAIN BLOCK ---------------------------------------
VARs = varSettings()
VARs.Read_Input()

# --- Logic to skip calculation on worksheet which has 'NO' keyword at the end and crunch stats for each sheet. -----
for vInvSheet in VARs.workbook.sheets():
	if vInvSheet.name.find("NO") != -1:
		print("Skipping sheet - " , vInvSheet.name)
        
	else:
		Testworksheet = VARs.workbook.sheet_by_name(vInvSheet.name)

	# ----- Start of Sorting entire excel based on Date column -----
	    # Sorting of any column, just give the column number to target_column variable
		target_column = 0 # Sort on Date field
		# Convert the data into lowercase for easy Manipultation
		data = [[ icell.lower() if isinstance(icell, unicode) else icell for icell in Testworksheet.row_values(i)] for i in range(Testworksheet.nrows)] # returns list with rows with a list of columns
		#Header_labels = data[0]  # Header row
		data = data[1:]  # Complete data except the header
		data.sort(key = itemgetter(target_column))  # Complete sorted data in the variable 'data' based on the 'Date' column
	# ----- End of Sorting -----
				
		(MonthlyTagValues, TotalTagValues, TotalMonthlyExpense, deduct_AMT, D_monthly_dedcut) = Crunch_Data(data)  # Replacing original code to crunch all tags

		write_EXCEL_Report(vInvSheet, MonthlyTagValues, TotalTagValues, TotalMonthlyExpense, deduct_AMT, D_monthly_dedcut)

		VARs.D_grpTagReport = {}  # Reset_Variables

save_filePath = os.path.join(VARs.SCRIPT_PATH, "Exp_Calc_Report_")
VARs.wb.save(save_filePath + VARs.year + '.xls')

write_HTML_Report()
# ---------------------------- END OF MAIN BLOCK ---------------------------------------