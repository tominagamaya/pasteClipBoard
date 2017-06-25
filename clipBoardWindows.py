import win32clipboard as cb
import xlrd as xl

def read_workbook():
    # 任意のファイルを指定する
    book = xl.open_workbook(filename=u'C:\\python\\sample.xlsx')
    sheet_num = book.nsheets
    sheet = book.sheet_by_index(0)

    row_list = []
    for row_index in range(sheet.nrows):
        # 任意のExcel列を指定する
        row_list.append(sheet.cell_value(rowx=row_index, colx=0))

    return row_list

def convertListToText(textList):
    text = "\n".join(textList)
    return text

def set_clipboard(text):
    cb.OpenClipboard()
    cb.SetClipboardText(text, cb.CF_UNICODETEXT)
    cb.CloseClipboard()

# 現段階では使っていない
def get_clipboard():
    cb.OpenClipboard()
    data = cb.GetClipboardData()
    cb.CloseClipboard()
    return data

# クリップボードに指定ファイルの値を貼りつける
set_clipboard(convertListToText(read_workbook()))
print("ok copy!")
