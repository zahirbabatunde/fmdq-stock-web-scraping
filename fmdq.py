import time
import schedule
import smtplib
import ssl
import os
import datetime as dt
import calendar
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
from bs4 import BeautifulSoup as bsoup

# clear all schedule
schedule.clear()

# constants
my_url = 'https://www.fmdqgroup.com/'
sender_email = 'zythonbot@gmail.com'
receiver_email = 'babs1160@gmail.com'
password_email = os.environ.get('ZYTHONBOT_PASSWORD')
schedule_time_daily = '18:00'
schedule_time_weekly = '19:00'
schedule_time_monthly = '20:00'


def main():
    check_soup = [[0]]
    try:
        # instantiate chrome driver and get page source code
        while (len(check_soup[0]) == 1) or (len(check_soup[0]) == 3):
            driver = webdriver.Chrome()
            driver.get(my_url)
            soup_fmdq = bsoup(driver.page_source, 'lxml')
            check_soup = list(soup_fmdq)
            if len(check_soup[0]) == 3:
                print('[Error: No internet connection]')
            driver.quit()
            time.sleep(10)

        date_of_quote = quotation_date(soup_fmdq)

        # find()- finds the particular element according to filter
        # find_all()- finds all the element in the previously specified filter
        # all the elements found are converted to list so each element is extracted
        grab_2027 = list(soup_fmdq.find('tr', {'class': "16.2884"}).find_all())
        grab_2037 = list(soup_fmdq.find('tr', {'class': "16.2499"}).find_all())
        grab_2049 = list(soup_fmdq.find('tr', {'class': "14.80"}).find_all())
        grab_2050 = list(soup_fmdq.find('tr', {'class': "12.98"}).find_all())

        bond_2027 = grab_data(grab_2027)
        bond_2037 = grab_data(grab_2037)
        bond_2049 = grab_data(grab_2049)
        bond_2050 = grab_data(grab_2050)

        dir_list = os.listdir()
        
        if 'fmdq.csv' not in dir_list:
            data = {'Date': [date_of_quote],
                    '17-MAR-2027': [float(bond_2027[1])],
                    '18-APR-2037': [float(bond_2037[1])],
                    '26-APR-2049': [float(bond_2049[1])],
                    '27-MAR-2050': [float(bond_2050[1])]}

            df_fmdq = pd.DataFrame(data)
            df_fmdq.to_csv('fmdq.csv', index=False)
            
        df_quote = pd.read_csv('fmdq.csv')
        last_quote = df_quote.iloc[-1, 0]
        new_row = {'Date': date_of_quote,
                   '17-MAR-2027': float(bond_2027[1]),
                   '18-APR-2037': float(bond_2037[1]),
                   '26-APR-2049': float(bond_2049[1]),
                   '27-MAR-2050': float(bond_2050[1])}

        if last_quote == new_row['Date']:
            pass
        else:
            df_quote = df_quote.append(new_row, ignore_index=True)

        df_quote.to_csv('fmdq.csv', index=False)

        text = '''\
        {}
        
        1.  Description: {}
            Price: {}
            Yield: {}
            Change: {}

        2.  Description: {}
            Price: {}
            Yield: {}
            Change: {}

        3.  Description: {}
            Price: {}
            Yield: {}
            Change: {}

        4.  Description: {}
            Price: {}
            Yield: {}
            Change: {}'''.format(date_of_quote,
                                 bond_2027[0], bond_2027[1], bond_2027[2], bond_2027[3],
                                 bond_2037[0], bond_2037[1], bond_2037[2], bond_2037[3], 
                                 bond_2049[0], bond_2049[1], bond_2049[2], bond_2049[3],
                                 bond_2050[0], bond_2050[1], bond_2050[2], bond_2050[3])

        html = """\
        <!DOCTYPE html>
        <html>
           <head>
              <style>
                 table, th, td {
                    border-collapse: collapse;
                    height: 20px;
                  }
                 table {
                    width: 50%%;
                    border: 1px solid #4B0082;
                 }
                 th, td{
                    text-align: center;
                    padding: 5px;
                    border: none;
                 }
                 th{
                    border-bottom: 1px solid #000;
                    background-color: #4B0082;
                    color: white;
                 }
                 td{
                    vertical-align: middle;
                 }
                 tr:nth-child(even) {
                  background-color: #f2f2f2;
                 }
                 tr{
                    border-bottom: 1px solid #000;
                 }
              </style>
           </head>
        
           <body>
              <table>
                <style>
                  caption{
                    text-align: center;
                    font-size: large;
                  }
                </style>
                <caption>%s</caption>
                 <tr>
                    <th>%s</th>
                    <th>%s</th>
                    <th>%s</th>
                    <th>%s</th>
                 </tr>
                 <tr>
                    <td>%s</td>
                    <td>%s</td>
                    <td>%s</td>
                    <td>%s</td>
                 </tr>
                 <tr>
                    <td>%s</td>
                    <td>%s</td>
                    <td>%s</td>
                    <td>%s</td>
                 </tr>
                 <tr>
                    <td>%s</td>
                    <td>%s</td>
                    <td>%s</td>
                    <td>%s</td>
                 </tr>
                 <tr>
                    <td>%s</td>
                    <td>%s</td>
                    <td>%s</td>
                    <td>%s</td>
                 </tr> 
              </table>
           </body>
        </html>
        """ % (date_of_quote, 'Description', 'Price', 'Yield', 'Change', 
               bond_2027[0], bond_2027[1], bond_2027[2], bond_2027[3],
               bond_2037[0], bond_2037[1], bond_2037[2], bond_2037[3],
               bond_2049[0], bond_2049[1], bond_2049[2], bond_2049[3],
               bond_2050[0], bond_2050[1], bond_2050[2], bond_2050[3])

        snd_mail(sender_email, receiver_email, password_email,
                 'FMDQ DAILY QUOTATION', text, html)

    except NoSuchElementException:
        print('error')
        time.sleep(2)


def monthly_quotation():

    mnthly_text = '''\
    Monthly Quote
    {}
    '''

    mnthly_html = '''\
    <html>
        <head></head>
        <body>
            <b>Monthly Quote</b>
            <br>
            {}
        </body>
    </html>
    '''

    if every_month():
        df_fmdq_mnth = pd.read_csv('fmdq.csv').tail(20)
        snd_mail(sender_email, receiver_email, password_email,
                 'FMDQ MONTHLY QUOTATION', mnthly_text, mnthly_html, True, df_fmdq_mnth)


def weekly_quotation():
    weekly_text = '''\
    Weekly Quote
    {}
    '''

    weekly_html = '''\
    <html>
        <head></head>
        <body>
            <b>Weekly Quote</b>
            <br>
            {}
        </body>
    </html>
    '''

    df_fmdq_week = pd.read_csv('fmdq.csv').tail()
    snd_mail(sender_email, receiver_email, password_email,
             'FMDQ WEEKLY QUOTATION', weekly_text, weekly_html, True, df_fmdq_week)


def quotation_date(parsed_html):
    quote_date = (parsed_html.find('tbody', {'class': 'tab item-9'})
                  .find('p').text.strip())
    return quote_date


def grab_data(grabbed):
    grab_list = []
    for td in grabbed:
        data = td.text.strip()
        grab_list.append(data)
    return grab_list


def every_month():
    today = dt.datetime.now()
    get_year = int(today.strftime('%Y'))
    get_month = int(today.strftime('%m'))
    get_today = int(today.strftime('%d'))

    days_in_month = calendar.monthrange(get_year, get_month)[1]

    if get_today >= days_in_month:
        return True
    else:
        return False


def snd_mail(snd_sender, snd_receiver, snd_password, subject,
             snd_text, snd_html, dataframe=False, df_file=None):
    snd_port = 465
    snd_server = 'smtp.gmail.com'

    mime_message = MIMEMultipart('alternative')
    mime_message['Subject'] = subject
    mime_message['From'] = snd_sender
    mime_message['To'] = snd_receiver

    if dataframe:
        snd_text = snd_text.format(df_file.to_string())
        snd_html = snd_html.format(df_file.to_html())

    snd_part1 = MIMEText(snd_text, 'text')
    snd_part2 = MIMEText(snd_html, 'html')

    mime_message.attach(snd_part1)
    mime_message.attach(snd_part2)

    snd_context = ssl.create_default_context()

    while True:
        try:
            with smtplib.SMTP_SSL(snd_server, snd_port, context=snd_context) as snd_server:
                snd_server.login(snd_sender, snd_password)
                snd_server.sendmail(snd_sender, snd_receiver, mime_message.as_string())
                print('Mail sent successfully')
        except Exception as e:
            print(e, '\n Retrying...')
            time.sleep(30)
            continue
        break


schedule.every().monday.at(schedule_time_daily).do(main)
schedule.every().tuesday.at(schedule_time_daily).do(main)
schedule.every().wednesday.at(schedule_time_daily).do(main)
schedule.every().thursday.at(schedule_time_daily).do(main)
schedule.every().friday.at(schedule_time_daily).do(main)
schedule.every().saturday.at(schedule_time_weekly).do(weekly_quotation)
schedule.every().day.at(schedule_time_monthly).do(monthly_quotation)

while True:
    schedule.run_pending()
    time.sleep(30)
