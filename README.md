# Automated Excel Report

## Description

A python program created for Finansia Syrus. This program aims to fill in a spread sheet with data related to Finansia Syrus. Crontab (not shown in github) only runs main.py every day including weekends.

## Python file features

### scan.py

Extracts data from the company's SQL database, output files and subprocesses. Regex is used to extract the data.

### selen.py

Automate the process of logging in for getting Finansia's Google Play and App Store Finansia's app information using Selenium. 

### mail.py

For sending the complete automated excel file to the Finansia senior staff. When there is an error in the program it sends a simple mail with an error message contained in it.

### main.py

Imports the functions from scan.py, selen.py and mail.py and arrange them for the tasks to work in a sequential manner. 
