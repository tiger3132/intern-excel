#Importing python classes created to use their functions
import scan
import selen
import finansia_hero
import mail

#For how long the program takes
import time

#For obtaining the exception
import traceback

#Importing operating system
from subprocess import run

#Import date
from datetime import datetime

start_time = time.time()
try:
    scan.check_day()
    sheet1, sheet2, wb_obj = scan.excel_create()
    row = scan.new_row(sheet1)
    print(row)
    scan.add_date(sheet1, row)
    
    scan.total_connect_user_count(sheet1, sheet2, row)
    scan.concurrent_user(sheet1, sheet2, row)
    scan.user_order_submission(sheet1, sheet2, row)
    scan.auto_user(sheet1, sheet2, row)
    scan.user_count_HTS_mode(sheet1, row)
    scan.download_count(sheet1, sheet2, row)
    scan.smart_login(sheet1, row)
    scan.user_connect_count(sheet2, row)
    #scan.matched_amount(sheet1, sheet2, row)
    scan.calculate_sum(sheet1, row)
    
    selen.google(sheet1, sheet2, row)
    selen.apple(sheet1, sheet2, row)

    #finansia_hero.full_process(sheet1, sheet2, row)
 
    scan.save_file(wb_obj)

    #mail.send_attachment()
except:
    #Check for errors for the whole program
    e = traceback.format_exc()
    print('Failed to do something: {}'.format(e))
    mail.send_error(e)

print(f'{time.time() - start_time:.2f}s')
