# -*- coding: utf-8 -*-
import xlrd

# This is for indicating a map between a column in dab and a column in excel. for example here
# the "full_name" is a column in the database and the "name" is a column field in excel structure should be like this:
# dynamic_excel_column = {
#     "db_column_name": {"name": "excel_column_name", "index": 0, "value": ""}, ...
# }

dynamic_excel_column = {
    "full_name": {"name": "name", "index": 0, "value": ""},
    "mobile": {"name": "Mobile", "index": 0, "value": ""},
}


def is_number(string):
    try:
        float(string)
        return True
    except ValueError:
        return False


def write_to_db(map, model_name):
    boiled_down_dynamic_excel_column = {}
    for key, value in dynamic_excel_column.items():
        # Here you can use some condition. For example when you want to put a birthday you should convert it to Datatime and
        # Here is the best place that you can do this sort of things.
        boiled_down_dynamic_excel_column.update({key: value["value"]})
    try:
        model_name.objects.create(**boiled_down_dynamic_excel_column)
    except Exception as e:
        print (boiled_down_dynamic_excel_column)
        print (e)


def calibrate_db_and_excel_columns(header_row_num, sheet):
    for i in range(sheet.ncols):
        excel_column_name = sheet.cell_value(header_row_num, i)
        for key, value in dynamic_excel_column.items():
            if excel_column_name == value['name']:
                value['index'] = i
                break


def read_excel(loc, header_row_num):
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)
    dict_row = {}
    calibrate_db_and_excel_columns(header_row_num, sheet)

    for i in range(header_row_num + 1, sheet.nrows):
        for key, value in dynamic_excel_column.items():
            value = sheet.cell_value(i, value['index'])
            if is_number(value):
                dynamic_excel_column[key]['value'] = int(value)
            else:
                dynamic_excel_column[key]['value'] = value

        write_to_db(dynamic_excel_column, Doctor)


if __name__ == "__main__":
    # Should start from zero.
    read_excel('PATH TO xlsx FILE', 0) 