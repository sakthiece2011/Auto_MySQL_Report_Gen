# Auto_MySQL_Report_Gen
A simple tool for generating a Excel spreadsheet(Dashboard) from MySql db

Requirement
--------------------------------
•	Python3 –windows Installer
•	Mysql Server, MySQL Windows Connector
•	Python Packages – mysqlclient, xlsxwriter, xlwt

Installation
--------------------------------
### Python
Step 1- Install Python 3.6 Windows installer (62 bit) -https://www.python.org/ftp/python/3.6.1/python-3.6.1-amd64.exe
Refer -All other Version of Python available in https://www.python.org/downloads/windows/
Step 2: Installation guide for python and pip - Please follow the instruction in http://matthewhorne.me/how-to-install-python-and-pip-on-windows-10/

### Python Packages:
Step 3: Run the pip commands --  pip install mysqlclient, pip install pymysql
Step 4: Run the pip command for excel python package - pip install xlsxwriter ,pip install xlwt

### Run
Step 5: Run the Python script – python IDBI_DashBoard.py(attached)
Step 6: Excel file will be generated(find in repo sample) in the same script path where it got executed.

### Need to do
```
1.	Change the database connection in the python script under (# Open database connection )by giving the host,username,password and Database name based on server we need to fetch the records.
```
