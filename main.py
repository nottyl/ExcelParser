#!/usr/bin/env python3

import excelparser as parse

# Initialize by parsing the needed data
parse.init('input_file.xlsx')

# Start reformatting the data of the parsed_file.xlsx
parse.filter_and_sort('parsed_file.xlsx')

# Start splitting the categories into different sheets
parse.split_categories('output_file.xlsx')
parse.sort_categories('output_file.xlsx')

# Formatting the spreadsheets