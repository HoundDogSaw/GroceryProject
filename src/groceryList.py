#!/usr/bin/env python
import openpyxl
from functions import build_header, build_list, build_menu
from constants import produce_list, meat_list, dairy_list
from constants import frozen_list, gluten_free_list, dry_food_list


def main():
    # This creates a blank workbook
    wb = openpyxl.Workbook()
    # Sets the current worksheet
    ws = wb.active
    # Setting columns width
    ws.column_dimensions['A'].width = 13.1
    ws.column_dimensions['D'].width = 13.1
    ws.column_dimensions['G'].width = 13.1
    ws.column_dimensions['B'].width = 2.7
    ws.column_dimensions['E'].width = 2.7
    ws.column_dimensions['H'].width = 2.7
    ws.column_dimensions['C'].width = 0.5
    ws.column_dimensions['F'].width = 0.5
    ws.column_dimensions['I'].width = 0.5

    ing_needed = build_menu()
    # creating the headers
    build_header(ws, 'Produce')
    build_list(ws, ing_needed, produce_list)
    build_header(ws, 'Meat')
    build_list(ws, ing_needed, meat_list)
    build_header(ws, 'Dairy')
    build_list(ws, ing_needed, dairy_list)
    build_header(ws, 'Frozen')
    build_list(ws, ing_needed, frozen_list)
    build_header(ws, 'Gluten Free')
    build_list(ws, ing_needed, gluten_free_list)
    build_header(ws, 'Dry Foods')
    build_list(ws, ing_needed, dry_food_list)

    # Remove the first three empty line
    # ws.delete_rows(1, 4)
    # This will save the workbook. it will overwrite the
    # current file.
    wb.save('src/testGroceryList.xlsx')
    wb.close()
    return


if __name__ == "__main__":
    main()
