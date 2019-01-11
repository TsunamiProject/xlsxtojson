from openpyxl import load_workbook
import json
from datetime import date, datetime


def json_serial(obj):
    if isinstance(obj, (datetime, date)):
        return obj.isoformat()


wb = load_workbook(filename='example.xlsx')

print(wb.get_sheet_names())

sheet = wb.active

main = {}
main['start_date'] = sheet.cell(row=2, column=3).value
main['end_date'] = sheet.cell(row=2, column=sheet.max_column).value

ll = {'FH': ['plane_name', 'company_name', 'statistic'],
      'Removals': ['block_number', 'manufacturer_company', 'number_of_deletions'],
      'Failures': ['block_number', 'manufacturer_company', 'confirmed_number_of_verified_faulty_blocks'],
      'Induced': ['block_number', 'manufacturer_company', 'number_of_forcedly_removed_blocks']}
for key in wb.get_sheet_names():
    main_dict = []
    for i in range(wb.get_sheet_by_name(key).min_row + 1, wb.get_sheet_by_name(key).max_row):
        buff_dict = {}
        stat_dict = []
        for j in range(wb.get_sheet_by_name(key).min_column + 2, wb.get_sheet_by_name(key).max_column):
            stat_dict.append(wb.get_sheet_by_name(key).cell(row=i, column=j).value)
        buff_dict[ll[key][0]] = wb.get_sheet_by_name(key).cell(row=i, column=1).value
        buff_dict[ll[key][1]] = wb.get_sheet_by_name(key).cell(row=i, column=2).value
        buff_dict[ll[key][2]] = stat_dict
        main_dict.append(buff_dict)
        main[key] = main_dict

with open('result.json', 'w') as fp:
    json.dump(main, fp, ensure_ascii=False, default=json_serial)
