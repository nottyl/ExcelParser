#!/usr/bin/env python3

import excelparser as parse

# Initialize by parsing the needed data
parse.init('input_file.xlsx')

# Start reformatting the data of the parsed_file.xlsx
parse.generate_new_file('parsed_file.xlsx')
