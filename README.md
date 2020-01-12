# Getting Excell pivot table from two incoming
The program implements reading data from two tables with suppliers and their payments1 (pos.xls and opl.xlsx, respectively), writing them to the MS Access database (Report.accdb) and generating, based on all the data in the database, a summary report (Report.xlsx) according to the rule: suppliers with data on the contract, and the amount of payments per month.
## Installation
The program is written in Python 3.7 and it is expected that it is already installed
In case of conflict with Python2, use pip3. Install the necessary modules:
```
  pip install -r requirements.txt
```
The program also uses tkinter, but since it comes preinstalled in most shipments, it is not specified. If it is not installed:
```
  pip install python3-tkinter
```
## Using
The main file is forms.py. For its regular work, it is necessary to have a database file (Report.accdb) in the same directory. The output Report.xlsx file is generated in the same directory.


Svod1.py - implementation of the same logic, but using only incoming files, and generating an output file without using a database.

## Project Description
Written for a specific configuration of Excel files and the data contained in them, the possibility of a different data structure is not taken into account.
