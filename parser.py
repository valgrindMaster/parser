from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

# Load spreadsheet for data processing.
wb = load_workbook('data.xlsx')

# Get spreadsheet.
ws = wb['Combined PowerGen CUAP and RS -']

# Mark all header starting indexes.
headerStartingIndexes = []

for cell in ws[1]:
  if cell.column < 8:
    continue
  if cell.value:
    headerStartingIndexes.append(cell.column)

# Create copy of worksheet for editing.
ws2 = wb.copy_worksheet(ws)
ws2.delete_rows(2)

# There are certain headers that actually utilize their full width of columns, so skip those.
headerExceptions = ['AL']
columnOffset = 0

headerStartingIndexesLen = len(headerStartingIndexes)

for i in range(headerStartingIndexesLen):
  print('~' + str(i // headerStartingIndexesLen), '% Complete.')
  colStart = headerStartingIndexes[i]
  colEnd = ws.max_column if (i == headerStartingIndexesLen - 1) else headerStartingIndexes[i + 1]

  if get_column_letter(colStart) in headerExceptions:
    continue

  columnOffsetUpdated = False

  for j in range(2, ws.max_row):
    combinedCellValues = set()
    for k in range(colStart, colEnd):
      value = ws.cell(row = j, column = k).value
      if value != None:
        value = str(value).strip().replace(r'^[\r\n\t]*\b|\b[\r\n\t]*$', '')
        combinedCellValues.add(value)
    if len(combinedCellValues) > 0:
      ws2.cell(row = j, column = colStart, value = (', '.join(combinedCellValues)))
      ws2.delete_cols(colStart - columnOffset, colEnd - columnOffset)
      if not columnOffsetUpdated:
        columnOffset += abs(colStart - colEnd) - 1
        columnOffsetUpdated = True

newWb = Workbook()
newWb.copy_worksheet(ws2)

newWb.save('temp.xlsx')