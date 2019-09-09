import os
import sys

if hasattr(sys, 'frozen'):
    os.environ['PATH'] = sys._MEIPASS + ";" + os.environ['PATH']

import time

from PyQt5.QtCore import *

import requests
import xlrd
from PyQt5.QtWidgets import QApplication, QWidget
from PyQt5.QtWidgets import (QLabel)
from PyQt5.QtWidgets import (QPushButton, QLineEdit)
from xlutils.copy import copy
from PyQt5.QtWidgets import QFileDialog


class myThread(QThread):  # 继承父类threading.Thread
    breakSignal = pyqtSignal(str)

    # def __init__(self, threadID, name, counter,pathname):
    #             #     threading.Thread.__init__(self)
    #             #     self.threadID = threadID
    #             #     self.name = name
    #             #     self.counter = counter
    #             #     self.pathname=pathname
    #             #     self.overall_len=0

    def __init__(self, pathname, roomid):
        #
        super(myThread, self).__init__()
        self.pathname = pathname
        self.overall_len = 0
        self.roomid = roomid


    def start1(self):
        self.start()


    def run(self):  # 把要执行的代码写到run函数里面 线程在创建后会直接运行run函数

        col = 0

        textcnt = 0
        textlen = 0
        # textlen=0
        temp=0
        # self.pathname+="danmu.xlsx"
        # # book = Workbook(encoding='utf-8')
        # book = Workbook()
        # sheet1 = book.add_sheet('Sheet 1')
        # book.save(self.pathname)

        work = xlrd.open_workbook(self.pathname)

        old_content = copy(work)
        ws = old_content.get_sheet(0)
        # print(ws)
        # ws.write(col, 1, "ssa")
        sheet = work.sheet_by_index(0)

        lasttext = []
        # lasttext = [0] * textlen
        lasttime = "1099-09-02 19:26:15"

        def comparetime(str1, str2):
            if int(str1[0:4]) > int(str2[0:4]):
                return 1
            elif int(str1[0:4]) < int(str2[0:4]):
                return 0

            if int(str1[5:7]) > int(str2[5:7]):
                return 1
            elif int(str1[5:7]) < int(str2[5:7]):
                return 0

            if int(str1[8:10]) > int(str2[8:10]):
                return 1
            elif int(str1[8:10]) < int(str2[8:10]):
                return 0

            if int(str1[11:13]) > int(str2[11:13]):
                return 1
            elif int(str1[11:13]) < int(str2[11:13]):
                return 0

            if int(str1[14:16]) > int(str2[14:16]):
                return 1
            elif int(str1[14:16]) < int(str2[14:16]):
                return 0

            if int(str1[17:19]) > int(str2[17:19]):
                return 1
            elif int(str1[17:19]) < int(str2[17:19]):
                return 0

            return -1

        while True:
            url = 'https://api.live.bilibili.com/ajax/msg'
            form = {
                "roomid": int(self.roomid),
                "csrf_token": "c54db958f52a40834dd438acc037702c"
            }
            # oneseconddanmu = []
            req = requests.post(url, data=form)
            html = req.content
            html_doc = str(html, 'utf-8')

            # print(html_doc)

            text_start = 0
            textlen = 0

            while True:
                text_start = html_doc.find("text", text_start)
                if text_start != -1:
                    text_start += 7
                else:
                    break
                # name_start=html_doc.find("nickname",name_start)
                # text_end=name_start-4
                # name_start+=11
                # name_end=html_doc.find("uname_color",name_start)-4
                textlen += 1

            colindex = max(col - textlen, 0)

            text_start = 0
            name_start = 0
            timestart = 0

            newtext_flag = 0

            print(html_doc)
            while True:
                ws.write(col, 1, "ssa")
                text_start = html_doc.find("text", text_start)
                if text_start != -1:
                    text_start += 7
                else:
                    break
                name_start = html_doc.find("nickname", name_start)
                text_end = name_start - 4
                name_start += 11
                name_end = html_doc.find("uname_color", name_start) - 4
                timestart = html_doc.find("timeline", text_start) + 11
                timeend = html_doc.find("isadmin", text_start) - 4
                danmutime = html_doc[timestart:timeend + 1]
                # print(danmutime)

                # if newtext_flag == 0:
                #     if html_doc[text_start:text_end + 1] not in lasttext:
                #         newtext_flag = 1
                # if newtext_flag == 1:
                # text.append(html_doc[text_start:text_end+1])
                # name.append(html_doc[name_start:name_end+1])
                # starttime = time.clock()]
                # if  (len(lasttext)==textlen and len(lasttext)>0):
                #     lasttext.pop(0)
                # if len(lasttext) == 1000:
                #     lasttext.pop(0)
                # lasttext.append(html_doc[text_start:text_end + 1])
                # print(lasttext)
                if comparetime(danmutime, lasttime) == 1:
                    print("name=%s danmutime=%s lasttime=%s" % (html_doc[name_start:name_end + 1], danmutime, lasttime))
                    lasttime = danmutime
                    ws.write(col, 0, html_doc[name_start:name_end + 1])
                    ws.write(col, 1, html_doc[text_start:text_end + 1])
                    ws.write(col, 2, danmutime)
                    col += 1
                    self.overall_len += 1
                    oneseconddanmu = []
                    oneseconddanmu.append(html_doc[name_start:name_end + 1])

                elif comparetime(danmutime, lasttime) == -1:
                    if html_doc[name_start:name_end + 1] not in oneseconddanmu:
                        ws.write(col, 0, html_doc[name_start:name_end + 1])
                        ws.write(col, 1, html_doc[text_start:text_end + 1])
                        ws.write(col, 2, danmutime)
                        oneseconddanmu.append(html_doc[name_start:name_end + 1])
                        print("name=%s danmutime=%s lasttime=%s" % (
                            html_doc[name_start:name_end + 1], danmutime, lasttime))
                        old_content.save("E:\danmu.xlsx")
                        col += 1
                        self.overall_len += 1

                    # endtime = time.clock()
                    # print((endtime-starttime))

                # lasttext.append(html_doc[text_start:text_end + 1])
                # textlen+=1
            textcnt += 1
            print(self.overall_len)
            if temp<self.overall_len:
                self.breakSignal.emit(str(self.overall_len))
            temp=self.overall_len
            # time.sleep(0.1)
            time.sleep(0.1)



class Example(QWidget):

    def __init__(self):
        super(Example,self).__init__()

        self.initUI()
        self.download_path = ""

        # self.lbl4=lbl4 = QLabel('0', self)

    def initUI(self):
        lbl1 = QLabel('存储文件夹位置:', self)
        lbl1.move(30, 50)

        lbl2 = QLabel('房间id:', self)
        lbl2.move(40, 80)

        lbl3 = QLabel('弹幕累计数目:', self)
        lbl3.move(60, 180)

        self.lbl4 = QLabel('0', self)
        self.lbl4.move(140, 180)

        # lbl1 = QLabel('csrf_token:', self)
        # lbl1.move(20, 60)
        # self.btn = QPushButton('Dialog', self)
        # self.btn.move(20, 20)
        # self.btn.clicked.connect(self.showDialog)

        self.le = QLineEdit(self)
        self.le.move(100, 80)
        # self.le1 = QLineEdit(self)
        # self.le1.move(100, 60)

        self.button = QPushButton('ok', self)
        self.button.move(100, 120)
        self.button.clicked.connect(self.on_click)

        self.button2 = QPushButton('选择', self)
        self.button2.move(130, 45)
        self.button2.clicked.connect(self.on_click1)

        self.setGeometry(300, 300, 300, 220)
        self.setWindowTitle('B站直播弹幕爬虫软件')
        # self.setWindowIcon(QIcon('web.png'))

        self.show()

    def changetext(self, a):
        self.lbl4.setText(a)
        self.lbl4.adjustSize()

    def on_click1(self):
        # QFileDialog.ge
        # self.download_path = QFileDialog.getExistingDirectory(self,"浏览","E:\workspace")
        # self.download_path += "danmu.xlsx"
        # book = Workbook(encoding='utf-8')
        # book = Workbook()
        # sheet1 = book.add_sheet('Sheet 1')
        # book.save(download_path)
        # print(self.download_path)
        self.download_path = QFileDialog.getOpenFileName(self, "浏览", "E:\workspace")[0]
        print(self.download_path)

    def on_click(self):
        textboxValue = self.le.text()
        roomid = ""
        roomid = textboxValue
        flag = 0

        self.thread1 = myThread(self.download_path, roomid)
        self.thread1.breakSignal.connect(self.changetext)
        self.thread1.start1()

        oneseconddanmu = []




        # thread1 = myThread(1, "Thread-1", 1, self.download_path)
        # thread1 = myThread(self.download_path, roomid)
        # thread1.breakSignal.connect(self.changetext)
        # thread1.start()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())
