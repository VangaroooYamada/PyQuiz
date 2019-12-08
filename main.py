from PIL import Image
import sys
import glob
import re
import time

import pyocr
import pyocr.builders
import openpyxl as px


tools = pyocr.get_available_tools()
if len(tools) == 0:
    print("No OCR tool found")
    sys.exit(1)

tool = tools[0]
print("Will use tool '%s'" % (tool.get_name()))


langs = tool.get_available_languages()
print("Available languages: %s" % ", ".join(langs))
lang = langs[1]
print("Will use lang '%s'" % (lang))


inames = glob.iglob('./imgs/*.png')
trans_dict = {
    u'①': '1',
    u'②': '2',
    u'③': '3',
    u'④': '4',
    u'⑤': '5',
    u'⑥': '6',
    u'⑦': '7',
    u'⑧': '8',
    u'⑨': '9',
    u'⑩': '10',
    u'⑪': '11',
    u'⑫': '12',
    u'⑬': '13',
    u'⑭': '14',
    u'⑮': '15',
    u'⑯': '16',
    u'⑰': '17',
    u'⑱': '18',
    u'⑲': '19',
    u'⑳': '20',
    '!': '！',
    '?': '？',
    '`': '「',
    ' ': '',
    '\n': '',
}
QLPath = './QuizList.xlsx'
quiz_list = px.load_workbook(QLPath)
ws = quiz_list.worksheets[-1]
MR = quiz_list.worksheets[0].max_row
BR = 2          # Brank Row's num
i_count = 0     # Processed Image Counter
q_count = 0     # Added Quiz Counter


while BR < MR+1:
    if not ws['B' + str(BR)].value is None:
        BR += 1
        continue
    break


start = time.time()

for img in inames:
    txt = tool.image_to_string(
        Image.open(img),
        lang='jpn',
        builder=pyocr.builders.TextBuilder(tesseract_layout=3)
    )

    txt = txt.translate(str.maketrans(trans_dict))
    for q in re.finditer(r'(問題正解率:\d+\%|Q\.)(.*?でしょう？)', txt):
        # print(q.groups()[1])
        if BR > MR:
            tmp_num = int(ws.title[:-1]) + 1000
            quiz_list.copy_worksheet(quiz_list.worksheets[0])
            ws = quiz_list.worksheets[-1]
            ws.title = str(tmp_num) + '-'
            BR = 2

        ws['B' + str(BR)].value = q.groups()[1]
        q_count += 1
        BR += 1

    i_count += 1

end_time = time.time() - start


for i in range(3):
    try:
        myfile = open('./QuizList.xlsx', 'r+')
    except IOError:
        print('*****************************************')
        print('Excel file is opened! Please close Excel.')
        print('({} times left)'.format(3 - i))
        print('*****************************************')
        key = input('Close Excel and press any key.')
    else:
        break
else:
    print('Processing Error!')
    quit(1)

quiz_list.save('./QuizList.xlsx')


print('''
**********************
SUCCESS
{} Photos Processed
{} Quiz Added
Process Time: {}s
**********************
'''.format(i_count, q_count, round(end_time, 3)))
