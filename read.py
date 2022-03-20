from openpyxl import load_workbook          #讀入xlsx文件
from openpyxl.utils import range_boundaries
import json                                 #
import pandas
from pandas import DataFrame

workbook = load_workbook(filename="KuangxYonh.xlsx")
sheet = workbook["KuangxYonh"]

# df = DataFrame(workbook.values)
dics = {}
qim = {}

for row in sheet.iter_rows(min_row=3, max_row=100, #8546,
                           min_col=1, max_col=13,
                           values_only=True):
    dic_id = row[0]
    dic = {
        "轄字": row[12],
        "聲符": row[1],
        "上古擬音": row[5],
        "聲母": row[6],
        "音節類型": row[3],
        "等": row[7],
        "開合": row[8],
        "韻母": row[9],
        "聲調": row[10],
    }
    dics[dic_id] = dic
    qim[row[12]] = row[5]

with open('data.json', 'w', encoding='utf-8') as f:
    json.dump(dics, f, ensure_ascii=False, indent=4)

with open('qim.json', 'w', encoding='utf-8') as f:
    json.dump(qim, f, ensure_ascii=False, indent=4)

list_of_strings = [ f'{key}\t{qim[key]}' for key in qim ]

with open('dangqskaaq.dict.yaml', 'w') as my_file:
    [ my_file.write(f'{st}\n') for st in list_of_strings ]

# with open('dangqskaaq.dict.yaml', 'w', encoding='utf-8') as f:
# file.write(json.dumps(exDict))
# for key, value in dic.items(): 
# f.write('%s:%s\n' % (key, value))
