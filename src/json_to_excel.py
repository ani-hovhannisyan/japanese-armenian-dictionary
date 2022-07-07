#!/usr/bin/python3
# Converting json file to excel file to sort visually 
# Dependencies: $ pip install pandas openpyxl json

import json
import pandas as pd
import os
from datetime import datetime

outputDir = "./dataset/"
timeNow = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")

dictFile = "./dataset/JAD_1.1. 初級(アルメニア語).xlsx"
print("Processable File is:", dictFile)

fileName = os.path.basename(dictFile).replace(".", "_")

sheet = "5課" # 1, 2, 3, 4, ...

outputFile = outputDir + fileName + "_" + sheet + "_" + timeNow + ".json"
print("New generatable file name will be:", fileName)

excelFile = pd.read_excel(dictFile, sheet_name=sheet)

parsed = json.loads(excelFile.to_json())

js = excelFile.to_json(orient='records')

labels = []
for i in parsed:
    pi = parsed[i]
    labels.append(i)
print("-- all labels", labels)

index = labels[0]
word = labels[1]
kana = labels[2]
hy = labels[3]
en = labels[4]

wordsDict = [ ]
for j in parsed[index]:
    i = parsed[index]
    print(type(parsed[hy][j]),)
    wordEl = {
        "id": parsed[index][j],
        "kana": parsed[kana][j],
        "word": parsed[word][j],
        "translate": {
            "hy": [  parsed[hy][j].split(",") if parsed[hy][j] else parsed[hy][j] ],
            "en": [ parsed[en][j] ]
        }
        #"example: {"hy": [], "en": []}
        #"suffixes: ["~x", "x~"] # where x is a suffix before or after word,
        # e.g. word + suffix = x~, suffix + word => x~
    }
    wordsDict.append(wordEl) 
    print("---new word is:", wordEl)

def generate_result_file(fileName, fileContent):
    f = open(fileName, "w", encoding="utf-8")
    content = json.dumps(fileContent, indent=2, ensure_ascii=False)
    f.write(content)
    f.close()

generate_result_file(outputFile, wordsDict)

print("Excel file generation succeeded")
