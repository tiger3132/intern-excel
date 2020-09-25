# Automated Excel Report

## Description

A python program created by myself in Finansia Syrus. Crontab (not shown in github) only runs main.py. 

## scan.py

Extracts data from the company's SQL database, output files and subprocesses. I frequently used regex to extract the data.

## selen.py

I automate the process of logging in to get Finansia's Google Play and App Store app information using Selenium. 

## mail.py

For sending an automated excel file to the Finansia senior staff. When there is an error in the program it sends a simple mail.

## main.py

Imports the functions from the other python files and ordered them. So instead of having to run and use crontab on three python files, only main.py runs. 
