from PIL import Image
import json
import pandas as pd
from datetime import datetime
import os

from openpyxl import load_workbook



def folder_make(file):
    wb = load_workbook(file)
    sheet= wb.worksheets[0]
    for i, row in enumerate(sheet.rows):
        if i == 0:
            continue
        name = row[0].value
        if not os.path.exists('./'+name):
            os.makedirs('./'+name)

    
        
file = input("폴더명만들 엑셀파일명을 입력해주세요.")
data = folder_make(file)


