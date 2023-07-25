import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

def get_data_in_exel(fileName) -> dict:
    path = f"{fileName}.xlsx"

    workBook: Workbook = openpyxl.load_workbook(path)
    workSheet: Worksheet = None
    tableData = {}
    answersList = []
    #[row 1..n][column 0..n]
    for workSheet in workBook:
        tableData[workSheet.title] = {}
        for row in workSheet.iter_rows(min_row=2,values_only=True):
            if (row[0], row[1], row[3], row[4]) == (None, None, None, None):
                break
            iter = 0
            articleKey: int = -1
            for value in row:
                if iter == 0:
                    articleKey = value if value != 'standart_answer' else -1
                    tableData[workSheet.title][articleKey] = {}
                    iter += 1
                    continue
                rowKeyValue = workSheet[1][iter].value
                if not rowKeyValue:
                    answersText = ';'.join(answersList)
                    tableData[workSheet.title][articleKey]['answers'] = answersText
                    break
                keyValue = rowKeyValue.split("|")[-1]
                answerValue = keyValue.split("_")[0]
                if answerValue == 'answer':
                    if answerValue:
                        answersList.append(value)
                        iter += 1
                        continue
                    continue
                tableData[workSheet.title][articleKey][keyValue] = value
                iter += 1
            else:
                answersText = ';'.join(answersList)
                tableData[workSheet.title][articleKey]['answers'] = answersText

    return tableData
