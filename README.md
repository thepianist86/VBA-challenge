# VBA-challenge
VBA Challenge for NW Data Analytics Boot Camp

VBA Script 'tickerTracker.vb' will loop through rows of ticker data in a spreadsheet, identifying a stock's opening price on the earliest day in the list and the closing price on the last day in the list.

The script will then generate a table which shows the difference between first opening and last closing (formatting the cell in red if the change is negative and green if the change is positive), percent change over the course of the year, and overall volume.

A second table will generate which finds the greatest % increase, the greatest % decrease, and the greatest total volume for the time period. At this time the script does not identify the associated ticker abbreviations for this data.

Data to be analyzed should be in individual worksheets by time period to be analyzed (e.g. first day of a year to last day of a year), and sorted first by ticker abbreviation then ascending by date.