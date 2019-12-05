from PIL import Image
import sys
import glob
import re
import time

import pyocr
import pyocr.builders
import openpyxl as px
import yaml


tool = None
lang = None


def set_ocr(tool_n=0, lang_n=1):
    global tool, lang

    tools = pyocr.get_available_tools()
    if len(tools) == 0:
        print("No OCR tool found")
        sys.exit(1)
    tool = tools[tool_n]
    print("Will use tool '%s'" % (tool.get_name()))

    langs = tool.get_available_languages()
    print("Available languages: %s" % ", ".join(langs))
    lang = langs[lang_n]
    print("Will use lang '%s'" % (lang))


def quiz_processor(img_list):
    global tool, lang

    QLPath = './QuizList.xlsx'
    quiz_list = px.load_workbook(QLPath)
    ws = quiz_list.worksheets[-1]
    MR = quiz_list.worksheets[0].max_row
    BR = 2          # Brank Row's num
    i_count = 0     # Processed Image Counter
    q_count = 0     # Added Quiz Counter

    while BR < MR+1:    # Search minimum blank cell
        if not ws['B' + str(BR)].value is None:
            BR += 1
            continue
        break

    with open('./trans_dict.yaml') as tdict:    # transration list import
        trans_dict = yaml.safe_load(tdict)

    for img in img_list:
        txt = tool.image_to_string(
            Image.open(img),
            # lang='jpn',
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

    quiz_list.save('./QuizList.xlsx')

    return i_count, q_count


if __name__ == '__main__':
    inames = glob.iglob('./imgs/*.png')

    set_ocr()

    start_time = time.time()
    i_num, q_num = quiz_processor(inames)
    end_time = time.time() - start_time

    print('''
    **********************
    SUCCESS
    {} Photos Processed
    {} Quiz Added
    Process Time: {}s
    **********************
    '''.format(i_num, q_num, round(end_time, 3)))
