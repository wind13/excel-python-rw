from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
from helper import find_in_range

ot_xlsx = '加班数据-9月.xlsx'
ot_sheet_name = 'Sheet1'
ot_staff_id_column = 'F'
ot_hours_column = 'N'
ot_start_date_column = 'O'
ot_end_date_column = 'P'
ot_start_row = 2

ot_wb = load_workbook(ot_xlsx)
ot_ws = ot_wb[ot_sheet_name]
ot_end_row = ot_ws.max_row + 1

# print(ot_end_row) # 178

wh_xlsx = '9月billing人力-01.xlsx'
wh_sheet_name = 'Sheet1'
wh_staff_id_column = 'D'
wh_start_row = 3
wh_holiday_color = 'FFC4BD97'

wh_wb = load_workbook(wh_xlsx)
wh_ws = wh_wb[wh_sheet_name]

# print(wh_ws.max_row) //378
# print(wh_ws.max_column) //45

wh_end_row = wh_ws.max_row + 1
wh_end_column = wh_ws.max_column + 1
wh_end_column_letter = get_column_letter(wh_ws.max_column)

for row in range(ot_start_row, ot_end_row):
  staff_id = ot_ws[ot_staff_id_column + str(row)].value
  hours = ot_ws[ot_hours_column + str(row)].value
  start_day = ot_ws[ot_start_date_column + str(row)].value
  end_day = ot_ws[ot_end_date_column + str(row)].value

  range_sid_from = wh_staff_id_column + str(wh_start_row) # C3
  range_sid_to = wh_staff_id_column + str(wh_end_row) # C379
  wh_staff_range = wh_ws[range_sid_from +  ':' + range_sid_to]
  staff_cell = find_in_range(staff_id, wh_staff_range)

  if staff_cell == None:
    print('这个 ' + str(staff_id) + ' 在人力表中没找到对应的行。')
    break

  range_day_from = 'A1'
  range_day_to = wh_end_column_letter + '1'
  wh_day_range = wh_ws[range_day_from +  ':' + range_day_to]
  day_cell = find_in_range(start_day, wh_day_range)

  # print(staff_cell.row)
  # print(day_cell.column_letter)
  sid_day_cell = wh_ws[day_cell.column_letter + str(staff_cell.row)]
  sid_day_value = sid_day_cell.value
  # print(day_cell.fill.fgColor)
  # print(sid_day_value)
  # print(hours)
  is_holiday = day_cell.fill.fgColor.rgb != None and day_cell.fill.fgColor.rgb == wh_holiday_color
  h = hours if is_holiday else (8 + hours)
  wh_ws[day_cell.column_letter + str(staff_cell.row)] = h
  comment = Comment('extended service ' + str(h) + 'h', ' ')
  wh_ws[day_cell.column_letter + str(staff_cell.row)].comment = comment

  # print(sid_day_cell.value)

wh_wb.save('ot_hours.xlsx')

# print(ot_ws[ot_end_date_column + str(ot_end_row)].column)
# print(ot_ws[ot_end_date_column + str(ot_end_row)].column_letter)

