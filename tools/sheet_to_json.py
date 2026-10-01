"""엑셀 근무표 → JSON(값+병합). tools/test.mjs 에 넘겨 실제 양식으로 시험할 때 사용.
   python3 tools/sheet_to_json.py 근무표.xlsx out.json"""
import json, sys, openpyxl
ws = openpyxl.load_workbook(sys.argv[1], data_only=True).worksheets[0]
vals = [[c.value for c in row] for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column)]
merged = [dict(row=m.min_row, col=m.min_col, nrows=m.max_row - m.min_row + 1, ncols=m.max_col - m.min_col + 1) for m in ws.merged_cells.ranges]
json.dump(dict(values=vals, merged=merged), open(sys.argv[2], 'w'), ensure_ascii=False, default=str)
