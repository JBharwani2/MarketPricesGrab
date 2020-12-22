# MarketPricesGrab
Market Prices Grab is an automated process that scrapes data from a financial data webpage and compiles it into a Microsoft Excel spreadsheet for further analysis. It is optimized to run once daily and skip days in which the market is closed such as weekends or holidays (the webpage has no new updates on these days).

Webpage: [Yahoo Finance Historical Data](https://finance.yahoo.com/quote/CPSS/history?p=CPSS)

# Imported Libraries
[Beautiful Soup](https://www.crummy.com/software/BeautifulSoup/)

[OpenPyXL](https://openpyxl.readthedocs.io/en/stable/#)

# Batch File - pricegrab.bat
@py.exe (user specific path to python file) %*
