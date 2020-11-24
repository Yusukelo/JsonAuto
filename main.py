import openpyxl
import json


#ここに日付を入力---------------------------------------------------------------------------------
day = "4/16"
#----------------------------------------------------------------------------------------------
#ここに感染者のエクセルのパスを記入------------------------------------------------------------------
wb = openpyxl.load_workbook('./感染者20200416.xlsx',data_only=True)
#----------------------------------------------------------------------------------------------
'''
使用方法
1.上に日付、パスを挿入
2.ファイル実行
3.ターミナルにて文字列をコピー
4.JSONファイルの挿入部にコピペ
5.Jsonファイルフォーマット
'''
wa = wb.active
values = []
for i in range(3, 53):
    values.append(wa.cell(row=i, column=3).value)

#print(values)

prefects = ["北海道", "青森", "岩手", "宮城", "秋田", "山形", "福島", "茨城", "栃木", "群馬", "埼玉", "千葉", "東京", "神奈川", "新潟", "富山", "石川", "福井", "山梨", "長野", "岐阜", "静岡", "愛知", "三重", "滋賀",
                "京都府", "大阪", "兵庫", "奈良", "和歌山", "鳥取", "島根", "岡山", "広島", "山口", "徳島", "香川", "愛媛", "高知", "福岡", "佐賀", "長崎", "熊本", "大分", "宮崎", "鹿児島", "沖縄", "クルーズ船", "チャーター", "職員"]

# 数値→[]変換
json_list = []
for number in values:
    txt = ""
    for loop in range(number):
        txt += "{}"
        if not loop == number-1:
            txt += ","
        else:
            pass
    json_list.append(txt)

key_value = zip(prefects, json_list)

# 文字列作成
l=0
context =''
context +='"'+ day +'" :{'
for prefect, json_list in zip(prefects, json_list):
    context += '"'+prefect+'": { "confirmed": [' + json_list + '] }'
    if not l == 49:
        context += ','
    else:
        pass
    l += 1
context += '}'

print(context)
