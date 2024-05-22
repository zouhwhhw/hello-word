import datetime
import xlrd
import requests

webhook = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=56133d7f-f3f1-49a6-9250-dee88e0b8071"


current_time=datetime.datetime.today()

today = current_time.__format__("%Y.%m.%d")

path = "tabll.xls"

open_list = xlrd.open_workbook(path)
sheet = open_list.sheet_by_index(0)
raw = sheet.nrows

pan_dun = False
for ro in range(1,raw):
    date_row = sheet.row_values(ro)
    user = date_row[0]
    quanxian = date_row[1]
    date_time = date_row[4]

    if date_time == today:
        pan_dun = True
        msg_data = {
            "msgtype": "text",
            "text": {
                "content": date_time+" "+user+quanxian+"权限"+"已到期,请及时处理@所有人"
            }
        }
        response = requests.post(webhook, json=msg_data)
if not pan_dun:
    msg_data = {
        "msgtype": "text",
        "text": {
            "content": today+"无权限到期，无需处理"
        }
    }
    response = requests.post(webhook, json=msg_data)










