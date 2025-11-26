# gspreadsheets-finance
Scripts for Google Spreadsheets that provide functions to retrieve the closing price of a stock (US + EU markets) for a given date.

## Installation
1) Go to a spreadsheet in [Google Spreadsheets](https://docs.google.com/spreadsheets/)
2) In the _Extensions_ menu, click on __Apps Script__.
3) Copy the two scripts into a new project.
4) If you are using the US market script (`polygon.js`), you will need to put your free [API key](https://massive.com/dashboard/signup) in the file.

## Usage
In the examples below, the cell `A1` contains a Date.
- __US market:__
```python
=POLY_HIST("SPY";A1)
```
- __EU market:__
```python
=EURONEXT_HIST("LU1681048804-XPAR";A1)
```
