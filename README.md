# gspreadsheets-finance
Google Spreadsheets scripts that provide functions to retrieve the closing price of a stock (US+EU markets) for a given date.

## Installation
1) Go to a spreadsheet in [Google Spreadsheets](https://docs.google.com/spreadsheets/).
2) In the _Extensions_ menu, click on __Apps Script__.
3) Copy the script(s) you want to use into a new project. You may need to execute each function once (so permissions can be granted).
4) If you are using the US market script (`polygon.gs`), you will need to put your free [API key](https://massive.com/dashboard/signup) in the file.

## Usage
In the examples below, the cell `A1` contains a Date.
- __US market__ (`polygon.gs`):
```python
=POLY_HIST("SPY";A1)
```
- __EU market__ (`euronext.gs`):
```python
=EURONEXT_HIST("LU1681048804-XPAR";A1)
```
