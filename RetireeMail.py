from sqlite3 import Cursor
import win32com.client
import datetime
import pymssql
import os

print("System을 시작합니다...")

outlook=win32com.client.Dispatch("Outlook.Application")
Txoutlook = outlook.CreateItem(0)

print("OutLook 연결 완료")

server = ''
database = ''
username = ''
password = ''
conn = pymssql.connect(server, username, password, database, charset='UTF-8')
cursor = conn.cursor()

print(database+ "DB 연결 완료")

Txoutlook.To = ""
Txoutlook.CC = ""

print("퇴사자 사번을  입력하세요 : ")
retireeNo = str(input())

sql = """
"""

cursor.execute(sql)

result = cursor.fetchone()
while result:
    if result[0] == retireeNo :
        retiree = result[1]
        retireeDept = result[2]
    result = cursor.fetchone()

conn.close()

print("인수인계자를 입력하세요 : ")
takeover = str(input())

dateArr = ['월요일', '화요일', '수요일', '목요일', '금요일', '토요일', '일요일']

Txoutlook.Subject = "[퇴사자 알림]" + str(datetime.datetime.today())[0:10] + "("+ str(dateArr[datetime.datetime.today().weekday()])[0:1] +") "+retiree

Txoutlook.HTMLBody = """
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>HTML !DOCTYPE declaration</title>
    <style>
        body{
            font-family : "맑은 고딕";
            font-size : 10pt;
        }

        .sub{color : rgb(192,0,0); font-weight : bold; font-size : 11pt; font-family : "맑은 고딕"}
        .a{color : black; font-weight : bold; font-size : 9pt; font-family : "맑은 고딕"}
        .jusungimg > img{
            float: left;
            width : 500px;
        }
        .boldfont{
            color: rgb(31, 73, 125);
            font-weight: bold;
        }
        .redfont{
            color: red;
            font-weight: bold;
        }
        .smallnormalfont{
            color: rgb(128, 128, 128);
            font-weight: bold;
            font-size: 8.5pt;
            float: left;
        }
        .redsmaill{
            color: red;
            font-weight: bold;
            font-size: 8.5pt;
        }
        .signboldfont{
            color: rgb(128, 128, 128);
            font-weight: bold;
            font-size: 9pt;
        }
        .signboldfontf{
            color: rgb(128, 128, 128);
            font-weight: bold;
            font-size: 9pt;
            float: left;
        }
        .signlink{
            color: rgb(0, 0, 204);
            font-weight: bold;
            font-size: 8.5pt;
            float: left;
            text-decoration: underline;
        }
        .sign{
            line-height : 85%;
        }
        .signnotice{
            color: red;
            font-weight: bold;
            font-size: 10pt;
        }
        .signnoticetext{
            color: rgb(31, 73, 125);
        }
    </style>

</head>


<body>

<div class="sub">
    <span class="a">◈</span> 1등 8단계 : <br>
    모범 ← 결정 ← 판단 ← 분석 ← 정리 ← 실험 ← 계획 ← 목표
</div>
<Br>
안녕하세요<br>
<br><br>

금일 퇴사자 정보입니다.<br><br>

이름 : """ + retiree + """<br>
사번 : """ + retireeNo + """<br>
부서 : """ + retireeDept + """<br>
전산비품 반납 : 완료<br>
DRM 계정 삭제 : 완료<br>
업무인수인계자 : """ + takeover +"""<br><br>

감사합니다.<br>
김성수 드림<br><br>


<br>


</body>
</html>


"""

# Txoutlook.Display(True)

Txoutlook.send

# os.system("pause")
