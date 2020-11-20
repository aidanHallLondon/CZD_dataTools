
# CZD data cleansing tools
# Cleanse tool – Python3
Author: Aidan Hall
November 2020 
https://github.com/aidanHallLondon/CZD/blob/master/cleanse.py

 loads data from a specific Spreadsheet file 
 (built in a very specific way) processes it and generates a new spreadsheet.

This tool takes the ReadingData spreadsheet and generates a new xlsx file with  the columns you specify it also adds binary columns for some dimensions (Columns).

 For those columns it adds a new column for all unique values and sets to 1/0 if there is a match
 Also adds a meta data sheet to help debug  the  data and code
