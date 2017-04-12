#!/usr/bin/python3

import MySQLdb as PyMySQL
from xlsxwriter.workbook import Workbook
import xlwt

# Open database connection
db = PyMySQL.connect("HOSTNAME","USERNAME","PASSWORD","DB_NAME" )
table = 'applicant_account_info' # table you want to save

#Excel
workbook = Workbook('Auto_IDBI_DASHBOARD.xlsx')
sheet = workbook.add_worksheet()
sheet.set_column(0, 0, 50)
sheet.set_column(1, 1, 15)
sheet.set_column(2, 2, 20)
sheet.set_column(3, 3, 42)
sheet.set_column(8, 8, 40)
sheet.set_column(4,4, 24)
sheet.set_column(5,5, 15)
sheet.set_column(6,6, 15)
sheet.set_column(9,11,15)

#sheet.set_row(2, 50) 
sheet.set_default_row(hide_unused_rows=True)
bold = workbook.add_format({'bold': True,'border': 1})
border=workbook.add_format({'border': 1})

format = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#D3D3D3', 'border': 1})
sheet.write(0,0,"Enrollment",format)
#sheet.add_table('A2:B6')


# prepare a cursor object using cursor() method
cursor = db.cursor()

##############################################################################
#Query 1
##############################################################################

# execute SQL query using execute() method.
cursor.execute( "SELECT count(ACCOUNT_STATUS) as Total_Enrollment FROM %s;" % table)
results_1 = cursor.fetchall()
######Getting the columns
sheet.write(1,0,"Total_Enrollment",border)

######Getting the Rows   

for row in results_1:
    for item in row: #i.e. for each field in that row
	    sheet.write(1,1,item,border)  #write excel cell from the cursor at row 1
		
		
##############################################################################
#Query 2
##############################################################################
	
cursor.execute( "SELECT count(1) FROM %s WHERE ACCOUNT_STATUS = 'Approved';" % table)
results_2 = cursor.fetchall()
######Getting the columns

sheet.write(2,0,"Total number of Approved enrolment",border)

######Getting the Rows   

for row in results_2:
    for item in row: #i.e. for each field in that row
	    sheet.write(2,1,item,border)  #write excel cell from the cursor at row 1

##############################################################################
#Query 3
##############################################################################
		
cursor.execute( "SELECT count(1) FROM %s WHERE ACCOUNT_STATUS = 'Rejected';" % table)
results_3 = cursor.fetchall()
######Getting the columns
sheet.write(3,0,"Total number of In Process enrolment",border)

######Getting the Rows   

for row in results_3:
    for item in row: #i.e. for each field in that row
	    sheet.write(3,1,item,border)  #write excel cell from the cursor at row 1

##############################################################################
#Query 4
##############################################################################
	
cursor.execute( "SELECT count(1) FROM %s WHERE ACCOUNT_STATUS = 'In Process';" % table)
results_4 = cursor.fetchall()

######Getting the columns
sheet.write(4,0,"Total number of account number generated InActive",border)

######Getting the Rows   

for row in results_4:
    for item in row: #i.e. for each field in that row
	    sheet.write(4,1,item,border)  #write excel cell from the cursor at row 1

##############################################################################
#Query 5
##############################################################################

		
cursor.execute( "SELECT count(1) FROM %s WHERE ACCOUNT_STATUS = 'Rejected';" % table)
results_5 = cursor.fetchall()
######Getting the columns
sheet.write(5,0,"Total number of account number generated InActive",border)

######Getting the Rows   

for row in results_5:
    for item in row: #i.e. for each field in that row
	    sheet.write(5,1,item,border)  #write excel cell from the cursor at row 1

##############################################################################
#Query 6
##############################################################################
	
cursor.execute( "SELECT count(1) FROM %s WHERE ACCOUNT_STATUS = 'Account Generated - InActive';" % table)
results_6 = cursor.fetchall()

######Getting the columns
sheet.write(6,0,"Total number of account number generated InActive",border)

######Getting the Rows   

for row in results_6:
    for item in row: #i.e. for each field in that row
	    sheet.write(6,1,item,border)  #write excel cell from the cursor at row 1
		

##############################################################################
#Query 7 -Loan status
##############################################################################
	

######Getting the columns
sheet.write(16,0,"Loan status by amount",format)
sheet.write(17,0,"Stages",bold)
sheet.write(17,1,"Total Group #",bold)
sheet.write(17,2,"Total Member #",bold)
sheet.write(17,3,"Total Amount #",bold)
sheet.write(18,0,"Total Loan Initiated #",bold)
sheet.write(19,0,"Total Loan In Process #",bold)
sheet.write(20,0,"Total loan Rejected #",bold)
sheet.write(21,0,"Total loan Disbursed #",bold)

######Getting the Rows   

cursor.execute( "SELECT COUNT(1) from (select group_loan_code FROM  cbo_group_member_loan_mapping GROUP BY group_loan_code)l;" )
results_7 = cursor.fetchall()

for row in results_7:
    for item in row: #i.e. for each field in that row
	    sheet.write(18,1,item,border)  #write excel cell from the cursor at row 1		
	
cursor.execute( "SELECT COUNT(1) from (select member_loan_code FROM  cbo_group_member_loan_mapping GROUP BY member_loan_code)l;")
results_8 = cursor.fetchall()


for row in results_8:
    for item in row: #i.e. for each field in that row
	    sheet.write(18,2,item,border)  #write excel cell from the cursor at row 1

cursor.execute( "SELECT  ifnull(SUM(requested_amount),0) FROM loan_request WHERE loan_id IN (SELECT loan_id FROM loan WHERE loan_code IN (SELECT DISTINCT group_loan_code FROM cbo_group_member_loan_mapping))");
results_9 = cursor.fetchall()

for row in results_9:
    for item in row: #i.e. for each field in that row
	    sheet.write(18,3,item,border)  #write excel cell from the cursor at row 1

cursor.execute( "SELECT COUNT(1) FROM loan_detail WHERE loan_id IN (SELECT loan_id FROM loan WHERE loan_code IN (SELECT DISTINCT group_loan_code FROM cbo_group_member_loan_mapping));");
results_10 = cursor.fetchall()

for row in results_10:
    for item in row: #i.e. for each field in that row
	    sheet.write(21,1,item,border)  #write excel cell from the cursor at row 1
		
cursor.execute( "SELECT COUNT(*) FROM loan_detail WHERE loan_id IN (SELECT loan_id FROM loan WHERE loan_code IN (SELECT DISTINCT member_loan_code FROM cbo_group_member_loan_mapping))");
results_11 = cursor.fetchall()

for row in results_11:
    for item in row: #i.e. for each field in that row
	    sheet.write(21,2,item,border)  #write excel cell from the cursor at row 1
		
cursor.execute( "SELECT ifnull(SUM(requested_amount),0) FROM loan_request WHERE loan_id IN (SELECT loan_id FROM loan_detail WHERE loan_id IN (SELECT loan_id FROM loan WHERE loan_code IN (SELECT DISTINCT group_loan_code FROM cbo_group_member_loan_mapping)))");
results_12 = cursor.fetchall()

for row in results_12:
    for item in row: #i.e. for each field in that row
	    sheet.write(21,3,item,border)  #write excel cell from the cursor at row 1

cursor.execute( "SELECT COUNT(*) FROM cbo_group_member_loan_mapping glc join loan lon on glc.group_loan_code = lon.loan_code and lon.status = 23")
results_13 = cursor.fetchall()

for row in results_13:
    for item in row: #i.e. for each field in that row
	    sheet.write(20,1,item,border)  #write excel cell from the cursor at row 1
		
cursor.execute( "SELECT COUNT(*) FROM cbo_group_member_loan_mapping glc join loan lon on glc.member_loan_code = lon.loan_code and lon.status = 23 ")
results_14 = cursor.fetchall()

for row in results_14:
    for item in row: #i.e. for each field in that row
	    sheet.write(20,2,item,border)  #write excel cell from the cursor at row 1

cursor.execute( "SELECT ifnull(SUM(requested_amount),0) FROM loan_request WHERE loan_id IN (select loan_id from loan where status = 23 and reference_id like '12%') ")
results_15 = cursor.fetchall()

for row in results_15:
    for item in row: #i.e. for each field in that row
	    sheet.write(20,3,item,border)  #write excel cell from the cursor at row 1
		
##############################################################################
#Query 7 -Loan
##############################################################################

sheet.write(0,3,"Loans",format)
sheet.write(1,4,"LoanCustomerMappingFile",bold)
sheet.write(1,5,"LoanJLGReqFile",bold)
sheet.write(2,3,"Total loan customer mapping requested #",border)
sheet.write(3,3,"waiting for bank response #",border)
sheet.write(4,3,"successfully processed #",border)
sheet.write(5,3,"Rejected #",border)

######Getting the Rows  

cursor.execute( "select count(1) from (SELECT jlg_id,idbi_cust_id,savings_account_number FROM loan_customer_bank_file_details)l")
results_16 = cursor.fetchall()

for row in results_16:
    for item in row: #i.e. for each field in that row
	    sheet.write(2,4,item,border)  #write excel cell from the cursor at row 1 
		
cursor.execute( "select count(1) from (SELECT jlg_id,idbi_cust_id,savings_account_number FROM loan_customer_bank_file_details WHERE request_status = 1)l")
results_17 = cursor.fetchall()

for row in results_17:
    for item in row: #i.e. for each field in that row
	    sheet.write(3,4,item,border)  #write excel cell from the cursor at row 1

cursor.execute( "select count(1) from (SELECT jlg_id,idbi_cust_id,savings_account_number FROM loan_customer_bank_file_details WHERE request_status = 2)l")
results_18 = cursor.fetchall()

for row in results_18:
    for item in row: #i.e. for each field in that row
	    sheet.write(4,4,item,border)  #write excel cell from the cursor at row 1

cursor.execute( "select count(1) from (SELECT jlg_id,idbi_cust_id,savings_account_number FROM loan_customer_bank_file_details WHERE request_status = 3)l")
results_19 = cursor.fetchall()

for row in results_19:
    for item in row: #i.e. for each field in that row
	    sheet.write(5,4,item,border)  #write excel cell from the cursor at row 1

cursor.execute( "select count(1) from (SELECT jlg_id,idbi_cust_id,product_code  FROM loan_request_bank_file_details)l")
results_20 = cursor.fetchall()

for row in results_20:
    for item in row: #i.e. for each field in that row
	    sheet.write(2,5,item,border)  #write excel cell from the cursor at row 1 

cursor.execute( "select count(1) from (SELECT jlg_id,idbi_cust_id,product_code  FROM loan_request_bank_file_details WHERE request_status = 1)l")
results_21 = cursor.fetchall()

for row in results_21:
    for item in row: #i.e. for each field in that row
	    sheet.write(3,5,item,border)  #write excel cell from the cursor at row 1 


cursor.execute( "select count(1) from (SELECT jlg_id,idbi_cust_id,product_code  FROM loan_request_bank_file_details WHERE request_status = 2)l")
results_22 = cursor.fetchall()

for row in results_22:
    for item in row: #i.e. for each field in that row
	    sheet.write(4,5,item,border)  #write excel cell from the cursor at row 1
		
cursor.execute( "select count(1) from (SELECT jlg_id,idbi_cust_id,product_code  FROM loan_request_bank_file_details WHERE request_status = 3)l")
results_23 = cursor.fetchall()

for row in results_23:
    for item in row: #i.e. for each field in that row
	    sheet.write(5,5,item,border)  #write excel cell from the cursor at row 1 
		
		
##############################################################################
#Query 7 -Stage
##############################################################################

sheet.write(10,0,"Loan",format)
sheet.write(10,1,"Requested",format)
sheet.write(10,2,"Waiting for Response",format)
sheet.write(10,3,"Success Response",format)
sheet.write(10,4,"Reject Response",format)
sheet.write(10,5,"Total",format)
sheet.write(11,0,"Stage_1",border)
sheet.write(12,0,"Stage_2",border)
sheet.write(13,0,"Stage_3",border)
sheet.write(14,0,"Stage_4",border)

######Getting the Rows  

cursor.execute( "SELECT COUNT(*) FROM loan_customer_bank_file_details WHERE request_status IN('1','2','3')")
results_24 = cursor.fetchall()

for row in results_24:
    for item in row: #i.e. for each field in that row
	    sheet.write(11,1,item,border)  #write excel cell from the cursor at row 1 
		
for row in results_24:
    for item in row: #i.e. for each field in that row
	    sheet.write(11,5,item,border)  #write excel cell from the cursor at row 1 
		
cursor.execute( "SELECT COUNT(*) FROM loan_customer_bank_file_details WHERE request_status IN('2')")
results_25 = cursor.fetchall()

for row in results_25:
    for item in row: #i.e. for each field in that row
	    sheet.write(11,2,item,border)  #write excel cell from the cursor at row 1 
		
cursor.execute( "SELECT COUNT(*) FROM loan_customer_bank_file_details WHERE request_status IN('1')")
results_26 = cursor.fetchall()

for row in results_26:
    for item in row: #i.e. for each field in that row
	    sheet.write(11,3,item,border)  #write excel cell from the cursor at row 1
		
cursor.execute( "SELECT COUNT(*) FROM loan_customer_bank_file_details WHERE request_status IN('3')")
results_27 = cursor.fetchall()

for row in results_27:
    for item in row: #i.e. for each field in that row
	    sheet.write(11,4,item,border)  #write excel cell from the cursor at row 1
		
cursor.execute( "SELECT COUNT(*) FROM loan_request_bank_file_details WHERE request_status IN('1','2','3')")
results_29 = cursor.fetchall()

for row in results_29:
    for item in row: #i.e. for each field in that row
	    sheet.write(12,1,item,border)  #write excel cell from the cursor at row 1 
		
for row in results_29:
    for item in row: #i.e. for each field in that row
	    sheet.write(12,5,item,border)  #write excel cell from the cursor at row 1

cursor.execute( "SELECT COUNT(*) FROM loan_request_bank_file_details WHERE request_status IN('2')")
results_30= cursor.fetchall()

for row in results_30:
    for item in row: #i.e. for each field in that row
	    sheet.write(12,2,item,border)  #write excel cell from the cursor at row 1 

cursor.execute( "SELECT COUNT(*) FROM loan_request_bank_file_details WHERE request_status IN('1')")
results_31= cursor.fetchall()

for row in results_31:
    for item in row: #i.e. for each field in that row
	    sheet.write(12,3,item,border)  #write excel cell from the cursor at row 1 
		
cursor.execute( "SELECT COUNT(*) FROM loan_request_bank_file_details WHERE request_status IN('3')")
results_32= cursor.fetchall()

for row in results_32:
    for item in row: #i.e. for each field in that row
	    sheet.write(12,4,item,border)  #write excel cell from the cursor at row 1 

cursor.execute( "SELECT COUNT(*) FROM loan_disbursement_bank_file_details WHERE request_status IN('1','2','3')")
results_33 = cursor.fetchall()

for row in results_33:
    for item in row: #i.e. for each field in that row
	    sheet.write(13,1,item,border)  #write excel cell from the cursor at row 1 
		
for row in results_33:
    for item in row: #i.e. for each field in that row
	    sheet.write(13,5,item,border)  #write excel cell from the cursor at row 1

cursor.execute( "SELECT COUNT(*) FROM loan_disbursement_bank_file_details WHERE request_status IN('2')")
results_34= cursor.fetchall()

for row in results_34:
    for item in row: #i.e. for each field in that row
	    sheet.write(13,2,item,border)  #write excel cell from the cursor at row 1 

cursor.execute( "SELECT COUNT(*) FROM loan_disbursement_bank_file_details WHERE request_status IN('1')")
results_35= cursor.fetchall()

for row in results_35:
    for item in row: #i.e. for each field in that row
	    sheet.write(13,3,item,border)  #write excel cell from the cursor at row 1 
		
cursor.execute( "SELECT COUNT(*) FROM loan_disbursement_bank_file_details WHERE request_status IN('3')")
results_36= cursor.fetchall()

for row in results_36:
    for item in row: #i.e. for each field in that row
	    sheet.write(13,4,item,border)  #write excel cell from the cursor at row 1

cursor.execute( "SELECT COUNT(*) FROM loan_repayment_bank_file_details WHERE request_status IN('1','2','3')")
results_37 = cursor.fetchall()

for row in results_37:
    for item in row: #i.e. for each field in that row
	    sheet.write(14,1,item,border)  #write excel cell from the cursor at row 1 
		
for row in results_37:
    for item in row: #i.e. for each field in that row
	    sheet.write(14,5,item,border)  #write excel cell from the cursor at row 1

cursor.execute( "SELECT COUNT(*) FROM loan_repayment_bank_file_details WHERE request_status IN('2')")
results_38= cursor.fetchall()

for row in results_38:
    for item in row: #i.e. for each field in that row
	    sheet.write(14,2,item,border)  #write excel cell from the cursor at row 1 

cursor.execute( "SELECT COUNT(*) FROM loan_repayment_bank_file_details WHERE request_status IN('1')")
results_39= cursor.fetchall()

for row in results_39:
    for item in row: #i.e. for each field in that row
	    sheet.write(14,3,item,border)  #write excel cell from the cursor at row 1 
		
cursor.execute( "SELECT COUNT(*) FROM loan_repayment_bank_file_details WHERE request_status IN('3')")
results_40= cursor.fetchall()

for row in results_40:
    for item in row: #i.e. for each field in that row
	    sheet.write(14,4,item,border)  #write excel cell from the cursor at row 1

		
cursor.close()
db.close()