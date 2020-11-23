
# CZD data cleansing tools

## Cleanse tool – Python3

| Author      | Aidan Hall                                                                                                                     |
| ----------- | ------------------------------------------------------------------------------------------------------------------------------ |
| last update | November 2020                                                                                                                  |
| Source      | [https://github.com/aidanHallLondon/CZD_dataTools](https://github.com/aidanHallLondon/CZD_dataTools) |
| license     | [MIT](LICENSE)                                                                                                                 |

# Overview

This tool takes the ReadingData spreadsheet and generates a new xlsx file with  the columns you specify and it can add binary columns for memebrs of dimensions (Columns) with a limited set of unique values.

 For those columns it adds a new column for all unique values and sets to 1/0 if there is a match. Also adds a meta data sheet to help debug the data and code.

 The same can be applied to delimeted lists of tokens. These usually need to be limited to the top few to avoid overwheming the output.

 It also creates a meta data sheet in to output  to help with debugging.

## Usage

`python3 cleanse.py`

> check or set the values in spreadsheetStructure. They should to point to the right xlsx source file and sheet as well as listiung all the data columns, their types and outputs. An ID column is needed to identify actual data rows and row 1 is always assumed to be the column names. A valid column is needed to further exlude bad rows of data.

    Execute python3 cleanse.py

 > Ignore errors about unrecognised extensions to the excel format.

    The output is another spreadsheet file.

 Cleanse loads data from a specific Spreadsheet file (built in a very specific way) and processes it and generating a new spreadsheet.

---

## Files

| File                         | Description                                                     |
| ---------------------------- | --------------------------------------------------------------- |
| cleanse.py                   | excutable code                                                  |
| spreadsheetStructure.py      | Settings and meta data about columns and how to process them    |
| xslsFormats.py               | Formatting for the output file                                  |
| LICENSE                      | License information                                             |
| README.md                    | Readme file in markdown/text                                    |
| Reading data to upload.xlsx  | anonimised source spreadsheet - not included to protect privacy |
| data.xlsx                    | Output file (not included)                                      |
| CZD_dataTools.code-workspace | VS Code workspace settings (not needed)                         |
