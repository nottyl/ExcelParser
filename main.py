#!/usr/bin/env python3
import excelparser as parse

# Initialize by parsing the needed data, please modify the 'columns' array into the columns you want to keep
columns = [2, 3, 5, 11, 25, 26, 28, 29, 31, 32, 34, 35, 37, 38, 40, 41, 43, 44, 46, 47]
parse.init('input_file.xlsx',  columns)

# Start reformatting the data of the parsed_file.xlsx
col_to_split = input("Enter the column that needs splitting: ")
parse.filter_and_sort('parsed_file.xlsx')
parse.split_column('output_file.xlsx', col_to_split)

# Start splitting the categories into different sheets
cat_to_split = input("Enter the category name to split: ")
parse.split_categories('output_file.xlsx', cat_to_split)
parse.sort_categories('output_file.xlsx', cat_to_split)

# Formatting the spreadsheets