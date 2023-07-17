import datetime
import json
import xlwt
import requests
import urllib.parse
from bs4 import BeautifulSoup
from urllib.request import urlopen

# 新建一个excel文件
workbook = xlwt.Workbook(encoding="utf-8")
# 新建一个sheet
worksheet = workbook.add_sheet("Problem")
# 表格样式
style = xlwt.XFStyle()  # 创建样式对象
alignment = xlwt.Alignment()  # 创建对齐方式对象
alignment.horz = xlwt.Alignment.HORZ_CENTER  # 设置水平方向居中对齐
alignment.vert = xlwt.Alignment.VERT_CENTER  # 设置垂直方向居中对齐
style.alignment = alignment  # 将对齐方式应用于样式对象
style.alignment.wrap = 1
# 写入表头
label = [
    "学号",
    "姓名",
    "uid",
    "通过题数",
    "暂无评定",
    "入门",
    "普及-",
    "普及/提高-",
    "普及+/提高",
    "提高+/省选-",
    "省选/NOI-",
    "NOI/NOI+/CTSC",
]
for i in range(len(label)):
    worksheet.write(0, i, label[i], style)
    worksheet.col(i).width = 3000
worksheet.col(11).width = 4000

err_id = []


def GetJson(luogu_id):
    headers = {
        "authority": "www.luogu.com.cn",
        "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "accept-language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
        "referer": "https://www.luogu.com.cn/",
        "sec-ch-ua": '"Chromium";v="116", "Not)A;Brand";v="24", "Microsoft Edge";v="116"',
        "sec-ch-ua-platform": '"Windows"',
        "sec-fetch-dest": "document",
        "sec-fetch-mode": "navigate",
        "sec-fetch-site": "same-origin",
        "sec-fetch-user": "?1",
        "upgrade-insecure-requests": "1",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36 Edg/116.0.0.0",
    }

    url = "https://www.luogu.com.cn/user/" + str(luogu_id)
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, "html.parser")
    return urllib.parse.unquote(soup.script.string.split('"')[1])


def array2str(array):
    str = ""
    for i in range(len(array)):
        str += array[i]
        if i != len(array) - 1:
            str += "\n"
    return str


def SaveXlsx(row, id, name, data):
    global workbook
    global worksheet

    # 写入数据
    worksheet.write(row, 0, id, style)
    worksheet.write(row, 1, name, style)
    worksheet.write(row, 2, data["currentData"]["user"]["uid"], style)
    worksheet.write(row, 3, data["currentData"]["user"]["passedProblemCount"], style)

    # 如果 data["currentData"]["passedProblems"] 不存在
    if "passedProblems" not in data["currentData"]:
        err_id.append(id + ":" + name)
        return

    Problems = []
    for _ in range(8):
        diff = []
        Problems.append(diff)

    for i in range(len(data["currentData"]["passedProblems"])):
        if data["currentData"]["passedProblems"][i]["difficulty"] == 0:
            Problems[0].append(data["currentData"]["passedProblems"][i]["pid"])
        elif data["currentData"]["passedProblems"][i]["difficulty"] == 1:
            Problems[1].append(data["currentData"]["passedProblems"][i]["pid"])
        elif data["currentData"]["passedProblems"][i]["difficulty"] == 2:
            Problems[2].append(data["currentData"]["passedProblems"][i]["pid"])
        elif data["currentData"]["passedProblems"][i]["difficulty"] == 3:
            Problems[3].append(data["currentData"]["passedProblems"][i]["pid"])
        elif data["currentData"]["passedProblems"][i]["difficulty"] == 4:
            Problems[4].append(data["currentData"]["passedProblems"][i]["pid"])
        elif data["currentData"]["passedProblems"][i]["difficulty"] == 5:
            Problems[5].append(data["currentData"]["passedProblems"][i]["pid"])
        elif data["currentData"]["passedProblems"][i]["difficulty"] == 6:
            Problems[6].append(data["currentData"]["passedProblems"][i]["pid"])
        elif data["currentData"]["passedProblems"][i]["difficulty"] == 7:
            Problems[7].append(data["currentData"]["passedProblems"][i]["pid"])

    worksheet.write(row, 4, array2str(Problems[0]), style)
    worksheet.write(row, 5, array2str(Problems[1]), style)
    worksheet.write(row, 6, array2str(Problems[2]), style)
    worksheet.write(row, 7, array2str(Problems[3]), style)
    worksheet.write(row, 8, array2str(Problems[4]), style)
    worksheet.write(row, 9, array2str(Problems[5]), style)
    worksheet.write(row, 10, array2str(Problems[6]), style)
    worksheet.write(row, 11, array2str(Problems[7]), style)

row = 1
with open("uid.txt", "r", encoding="utf-8") as f:
    for line in f:
        id, name, uid = line.split()
        data = json.loads(GetJson(uid))
        print(id, name, uid, data["currentData"]["user"]["passedProblemCount"])
        SaveXlsx(row, id, name, data)
        row += 1

print("err_id:", err_id)

# 保存表格
now = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
workbook.save("Luogu-" + now + ".xls")
