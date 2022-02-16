# Automated Excel Report

## Description

A python program created for Finansia Syrus. This program aims to fill in a spread sheet with data related to Finansia Syrus. Crontab (not shown in github) only runs main.py every day including weekends.

## Python file features

### scan.py

Creates the excel file. Write to the file the data extracted from the company's SQL database that are associated to customers' interaction inside Finansia's mobile and web application. Regex is used to extract the data needed from the whole string. 

### selen.py

Selenium is used to traverse Finansia's Google Play Insight and App Analytics website to obtain other data related to the customers' usage of the mobile app like Android download count. The login system have to be bypassed in order to reach the needed data. 

### mail.py

This sends an email with the completed excel file attached to it. When there is an error in the program a simple mail with an error message is sent instead.

### main.py

This python file is triggered by crontab. It first imports the functions from scan.py, selen.py and mail.py, and then arrange them sequentially so that individual tasks works at a particular period. For instance, the functions from scan.py and selen.py should be invoked before the send_attachment function from mail.py.  
