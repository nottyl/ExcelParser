#!/usr/bin/env python3
import excelparser as parse

# Initialize by parsing the needed data, please modify the 'columns' array into the columns you want to keep
columns = [2, 3, 5, 11, 25, 26, 28, 29, 31, 32, 34, 35, 37, 38, 40, 41, 43, 44, 46, 47]
parse.init('input_file.xlsx',  columns)

# Start reformatting the data of the parsed_file.xlsx
parse.filter_and_sort('parsed_file.xlsx')

# Start splitting the categories into different sheets
parse.split_categories('output_file.xlsx')
parse.sort_categories('output_file.xlsx')

# Formatting the spreadsheets
# TODO: Add formatting function (Borders, Merge cells, Font)