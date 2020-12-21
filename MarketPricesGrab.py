#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Daily Financial Data Scraper

This script allows the user to get financial data from Yahoo finance's historical data page. The data is then
compiled into an Excel spreadsheet where it can be used to compare specific trends and used to calculate
necessary information.

The program is optimized to run automatically once a day. Therefore it skips days in which the market is closed
and the webpage does not have a new update including weekends and holidays.

If the .xlsx file is open on the device when the script is scheduled to run, an error message will be sent to
remind the user to close the file. Also, if the file location is changed, the file_dir variable must be changed
to match the new file path.
"""

# --- Details --------------------------------------------------------------------------------------------------
__author__ = "Jeremy Bharwani"
__date__ = "12/17/2020"
__license__ = "MIT"
__maintainer__ = "Jeremy Bharwani"
__email__ = "jcb926@gmail.com"
__status__ = "Development"

# --- Built-ins ------------------------------------------------------------------------------------------------
import os
import datetime

# --- Other Libraries ------------------------------------------------------------------------------------------
import bs4
import requests
import openpyxl

# TODO: other formulas included each time it updates
# TODO: checks against other finance trackers to catch yahoo's possible errors


def main():
    yahoo_url = 'https://finance.yahoo.com/quote/CPSS/history?p=CPSS'
    file_name = 'PriceGrabTest.xlsx'
    file_dir = 'C:\\Users\\Jeremy\\Documents'  # file path not pushed to git
    titles = ['date', 'open', 'high', 'low', 'close', 'volume']

    data = scrape_data(yahoo_url, titles)
    convert_date(data)
    print_to_spreadsheet(data, titles, file_name, file_dir)


# --- Methods -------------------------------------------------------------------------------------------------

def scrape_data(url, titles):
    """
    Gets data from webpage and compiles it into a list that is returned to main. Six css selectors are used
    to find the positions where the latest day's data is stored on the website. These are retrieved by manually
    inspecting website's html script (liable to need update if html format is altered in the future)

    :param url: str
        Website url, can be changed in main if necessary
    :param titles: list
        List of each column's title, each in string format
    :return: dict
        Keys filled with column titles and webpage data inserted into those keys' values
    """
    data = {}

    # list of each css selector position, retrieved from inspecting website's html script
    soup_location = [r'#Col1-1-HistoricalDataTable-Proxy > section > div.Pb\(10px\).Ovx\(a\).W\(100\%\) > table > '
                     r'tbody > tr:nth-child(1) > td.Py\(10px\).Ta\(start\).Pend\(10px\) > span',
                     r'#Col1-1-HistoricalDataTable-Proxy > section > div.Pb\(10px\).Ovx\(a\).W\(100\%\) > table > '
                     r'tbody > tr:nth-child(1) > td:nth-child(2) > span',
                     r'#Col1-1-HistoricalDataTable-Proxy > section > div.Pb\(10px\).Ovx\(a\).W\(100\%\) > table > '
                     r'tbody > tr:nth-child(1) > td:nth-child(3) > span',
                     r'#Col1-1-HistoricalDataTable-Proxy > section > div.Pb\(10px\).Ovx\(a\).W\(100\%\) > table > '
                     r'tbody > tr:nth-child(1) > td:nth-child(4) > span',
                     r'#Col1-1-HistoricalDataTable-Proxy > section > div.Pb\(10px\).Ovx\(a\).W\(100\%\) > table > '
                     r'tbody > tr:nth-child(1) > td:nth-child(5) > span',
                     r'#Col1-1-HistoricalDataTable-Proxy > section > div.Pb\(10px\).Ovx\(a\).W\(100\%\) > table > '
                     r'tbody > tr:nth-child(1) > td:nth-child(7) > span']

    # accesses specified website's text
    webpage = requests.get(url)
    soup = bs4.BeautifulSoup(webpage.text, 'html.parser')

    # fill data dictionary with data from website
    for category in range(len(titles)):
        grab = soup.select(soup_location[category])
        data[titles[category]] = grab[0].text.strip()

    return data


def print_to_spreadsheet(data, titles, file_name, file_dir):
    """
    Opens the spreadsheet, receives the list of data and compiles it into the next empty row in the Excel
    spreadsheet, then saves the updated spreadsheet.

    :param data: dict
        Contains values scraped from website with keys matching the column titles which are found in
        the titles list
    :param titles: list
        List of each column's title, each in string format
    :param file_name: str
        Name of the Excel spreadsheet that is being updated, can be changed in main if necessary
    :param file_dir: str
        Directory in which the Excel spreadsheet is saved (overwrites previous save but should not
        delete anything, only adds to previous versions), can be changed in main if necessary
    """
    columns = ['A', 'B', 'C', 'D', 'E', 'F']
    confirm = True
    os.chdir(file_dir)

    # opens workbook if it is in the correct path and navigates to the next empty row
    try:
        workbook = openpyxl.load_workbook(file_name)
    except FileNotFoundError:
        print('Could not open file\nConfirm file path has not changed from:')
        print(os.getcwd())
        raise  # ends program with a possible error explanation
    sheet = workbook['Sheet1']
    row = next_empty_row(sheet)

    # inserts converted values into the spreadsheet
    for num in range(len(titles)):
        cell = columns[num] + str(row)
        if num == 0:
            check_market_open(data, sheet, row)
            sheet[cell].value = data['date']
            sheet[cell].number_format = u'mm/dd/yyyy'  # date format
        elif num == (len(titles) - 1):
            vol = data['volume']
            vol_num = vol.replace(',', '')
            sheet[cell].value = int(vol_num)
            sheet[cell].number_format = u'#,##0'  # volume format
        else:
            sheet[cell].value = float(data[titles[num]])
            sheet[cell].number_format = u'0.0000'  # format for all other values

    # saves workbook and avoids error message if file is still open
    try:
        workbook.save(file_name)
    except PermissionError:
        print('Could not update, the file must be closed on this device')
        confirm = False

    if confirm:
        print('File Update Confirmation')


def check_market_open(data, sheet, row):
    """
    Checks to see if the date being scraped from the webpage is the same as the last entry in the spreadsheet.
    If it is the same, the entire program is ended. This helps avoid duplicates and will not add extra entries
    on weekends or holidays when the markets are closed.

    :param data: dict
        Contains values scraped from website with keys matching the column titles which are found in
        the titles list
    :param sheet: dict
        Contains the details of every cell in the spreadsheet, can call .value to get the value within a
        specific cell
    :param row: int
        Row number of the first empty row from the top (must be converted into a string to be used to get
        the value from the cell in that row)
    """
    cell = 'A' + str(row - 1)
    if sheet[cell].value == data['date']:
        print('Markets are closed today. No update.')
        exit()


def next_empty_row(sheet):
    """
    Searches for the next empty row in the spreadsheet and returns it to the print_to_spreadsheet method.
    Starts at row 3 because above rows are used for titles and other info, not data.

    :param sheet: list
        Specified spreadsheet that this method is searching within
    :return: int
        Row number of the first empty row from the top
    """
    row = 3
    cell = 'A' + str(row)

    # checks each cell in column A until an empty one is found
    while sheet[cell].value is not None:
        row += 1
        cell = 'A' + str(row)
    return row


def convert_date(data):
    """
    Converts the string format date from the webpage into a list of three integer values which are then
    combined into a formatted date by the built-in datetime module. The new format is inserted back into
    the 'date' key of the data dictionary.

    :param data: dict
        Contains values scraped from website with keys matching the column titles which are found in
        the titles list
    """
    date = data['date']
    date_values = []

    month_num = datetime.datetime.strptime(date[0: 3], "%b")  # strips month
    date_values.append(int(date[8:12]))  # year
    date_values.append(month_num.month)  # month
    date_values.append(int(date[4:6]))  # day
    converted_date = datetime.datetime(date_values[0], date_values[1], date_values[2])
    data['date'] = converted_date


if __name__ == '__main__':
    main()
