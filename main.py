import pandas as pd
import openpyxl as op

read_csv = pd.read_csv('Levelling.csv')
read_csv.to_excel('level.xlsx')


work_book = op.load_workbook('level.xlsx')
ws = work_book.active

list_of_all_sights = []
list_of_Rise_n_Fall = []

for rows in ws.iter_rows(min_row=2, min_col=2, max_col=4):
    individual_sights = [cell.value for cell in rows if cell.value]
    list_of_all_sights.append(individual_sights)

print(list_of_all_sights)
for i, items in enumerate(list_of_all_sights):
    if i < (len(list_of_all_sights)-1):
        bs = list_of_all_sights[i][0]
        fs = list_of_all_sights[i+1][0]
        if len(list_of_all_sights[i]) == 2:
            bs = list_of_all_sights[i][0]
        if len(list_of_all_sights[i+1]) == 2:
            fs = list_of_all_sights[i+1][1]
        Rise_or_Fall = bs-fs
        list_of_Rise_n_Fall.append(Rise_or_Fall)

ws.insert_cols(5)
ws.insert_cols(6)
ws.delete_cols(1)
print(list_of_Rise_n_Fall)


work_book.save('results.xlsx')


work_book1 = op.load_workbook('results.xlsx')
ws1 = work_book1.active
BM = eval(input('Enter Your BenchMark'))

# -------------------------------------------------------------
# Computing for the Adjusted misclose per Angle in Cell D
ws1.cell(row=1, column=4).value = 'Rise/Fall'
for cols in ws1.iter_cols(min_col=4, min_row=3, max_col=4):
    # print(cols)
    for i, cells in enumerate(cols):
        # print(i)
        # print(cells.value)
        cells.value = list_of_Rise_n_Fall[i]


ws1.cell(row=1, column=5).value = 'RL'
for i in range(2, 3+len(list_of_Rise_n_Fall)):
    if i == 2:
        ws1.cell(row=i, column=5).value = BM
    else:
        results = float(ws1.cell(
            row=i-1, column=5).value) + float(ws1.cell(
                row=i, column=4).value)

        ws1.cell(row=i, column=5).value = results

        print(
            f'index={i}----rl={float(ws1.cell(row=i-1, column=5).value)} - rf={float(ws1.cell(row=i, column=4).value)} and results = {results} ')

work_book1.save('results.xlsx')
