#!/usr/bin/env python
# -*- coding: utf-8 -*-

import csv
import json
import re
import sys
from openpyxl import load_workbook

INDEX_WORKSHEET = "INDEX-filtered"
INDEX_WORKSHEET_VARIABLE_NAME_COLUMN = "2021 Mnemonic (variable)"


def index_worksheet_to_variable_list(index_worksheet):
    """
    To parse the index_worksheet into an object defining which worksheets in the workbook are to be processed, and what
    additional metadata is needed for each, first get worksheet_colnames, then use column names
    to get_variable_details from each row. Variables that do not have a hyperlink to a sheet in the workbook in the 
    INDEX_WORKSHEET_VARIABLE_NAME_COLUMN cannot be processed.
    """
    rows = list(index_worksheet.rows)
    colnames = [cell.value for cell in rows[0]]
    variable_list = [] 

    for non_header_row in rows[1:]:
    
        row_as_dict = dict(zip(colnames, non_header_row))
    
        variable_name_cell = row_as_dict[INDEX_WORKSHEET_VARIABLE_NAME_COLUMN]
        if variable_name_cell.hyperlink is None:
            print(f"Ignoring variable {variable_name_cell.value} as it does not link to a variable in spreadsheet.")
            continue
        else:
            variable_name = variable_name_from_hyperlink_cell(variable_name_cell)

        variable_details = { colname: cell.value for colname, cell in row_as_dict.items() }
        variable_details["variable_name"] = variable_name

        variable_list.append(variable_details)
    
    return variable_list
    

def variable_name_from_hyperlink_cell(hyperlink_cell):
    return hyperlink_cell.hyperlink.location.split('!')[0].upper().replace("'", "")
    
    


def main():
    workbook_filename = sys.argv[1]
    wb = load_workbook(workbook_filename)
    variable_list = index_worksheet_to_variable_list(wb[INDEX_WORKSHEET])
    print(variable_list)


if __name__ == "__main__":
    main()
