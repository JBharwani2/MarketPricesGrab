##########################################################################################################
##########################################################################################################
##########################################################################################################
# Program: MarketPricesGrab
# Author: Jeremy Bharwani
# Date Created: 12/17/20
# Date Last Updated: 12/18/20
# Description: This program scrapes data from yahoo finance's historical data page and compiles it into
#              an excel spreadsheet where it can be used to compare specific trends and used to calculate
#              necessary information. Can be run daily by pressing [WINDOWS-KEY + R] and typing
#              'pricegrab' into the run command box. Confirmation of successful data grab will be shown
#              in the command window that pops up.
# TODO: other formulas included each time it updates
# TODO: checks against other finance trackers to catch yahoo's possible errors
# TODO: automate to run once a day at a certain time
##########################################################################################################
##########################################################################################################
##########################################################################################################
import bs4, requests, openpyxl, os, datetime


def main():
    yahoo_url = 'https://finance.yahoo.com/quote/CPSS/history?p=CPSS'
    file_name = 'PriceGrabTest.xlsx'
    file_dir = 'C:\\...'  # file path not pushed to git
    titles = ['date', 'open', 'high', 'low', 'close', 'volume']

    data_list = scrape_data(yahoo_url, titles)
    print_to_spreadsheet(data_list, titles, file_name, file_dir)


##########################################################################################################
# Gets data from yahoo finance and compiles it into a list that is returned to main. Six css selectors
# are used to find the positions where the latest day's data is stored on the website. These are
# retrieved by manually inspecting website's html script (liable to need update if html format is altered
# in the future)
#
# url: website url, can be changed in main if necessary
# titles: list of each column's title, each in string format
##########################################################################################################
def scrape_data(url, titles):
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


##########################################################################################################
# Opens the spreadsheet, receives the list of data and compiles it into the next empty row in the Excel
# spreadsheet, then saves the updated spreadsheet.
#
# data: dictionary containing numbers scraped from website with keys matching the column titles which are
#       found in the titles list
# titles: list of each column's title, each in string format
# file_name: name of the Excel spreadsheet that is being updated, can be changed in main if necessary
# file_dir: directory in which the Excel spreadsheet is saved (overwrites previous save but should not
#           delete anything, only adds to previous versions), can be changed in main if necessary
##########################################################################################################
def print_to_spreadsheet(data, titles, file_name, file_dir):
    columns = ['A', 'B', 'C', 'D', 'E', 'F']
    confirm = True
    os.chdir(file_dir)

    # opens workbook's and navigates to the next empty row
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
            date = convert_date(data['date'])
            sheet[cell].value = datetime.datetime(date[0], date[1], date[2])
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


##########################################################################################################
# Searches for the next empty row in the spreadsheet and returns it to the print_to_spreadsheet method.
# Starts at row 3 because above rows are used for titles and other info, not data.
#
# sheet: specified spreadsheet that this method is searching within
##########################################################################################################
def next_empty_row(sheet):
    row = 3
    cell = 'A' + str(row)

    # checks each cell in column A until an empty one is found
    while sheet[cell].value is not None:
        row += 1
        cell = 'A' + str(row)
    return row


##########################################################################################################
# Converts the string format date from the website into a list of three integer values which is returned
# to the print_to_spreadsheet method to be input into spreadsheet with correct formatting.
# (list order: year, month, date)
#
# date: date in string format ("Jan 1, 2020")
##########################################################################################################
def convert_date(date):
    date_values = []

    month_num = datetime.datetime.strptime(date[0: 3], "%b")  # strips month
    date_values.append(int(date[8:12]))  # year
    date_values.append(month_num.month)  # month
    date_values.append(int(date[5:6]))  # day
    return date_values


if __name__ == '__main__':
    main()
