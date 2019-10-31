# -*- coding: UTF-8 -*-
import wx
import os
from shutil import copyfile
import xlrd
import pymysql

## 主界面
class MainFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, wx.ID_ANY, '导入文件', pos=wx.DefaultPosition,
                          size=(400, 400), style=wx.DEFAULT_FRAME_STYLE)
        self.CreateStatusBar()
        self.__BuildMenus()

        saveButton = wx.Button(self, label='导入', pos=(0, 5), size=(80, 25))
        label = wx.StaticText(self, -1, u"已选择文件")
        textBox = wx.TextCtrl(self, -1, style=wx.TE_MULTILINE, size=(-1,500))
        sizer = wx.BoxSizer(wx.HORIZONTAL)
        sizer.Add(label, 0, wx.ALL | wx.ALIGN_CENTRE)
        sizer.Add(textBox, 1, wx.ALL | wx.ALIGN_CENTRE)
        self.Bind(wx.EVT_BUTTON, self.OnClick, saveButton)
        saveButton.SetDefault()
        self.__TextBox = textBox
        self.SetSizerAndFit(sizer)
#创建菜单栏
    def __BuildMenus(self):
        mainMenuBar = wx.MenuBar()

        fileMenu = wx.Menu()

        fileMenuItem = fileMenu.Append(-1, "打开单个文件")
        self.Bind(wx.EVT_MENU, self.__OpenSingleFile, fileMenuItem)

        mainMenuBar.Append(fileMenu, title=u'&选择导入文件')

        self.SetMenuBar(mainMenuBar)

    def __OpenSingleFile(self, event):
        filesFilter = "Excel files (*.xls)|*.xlsx"
        fileDialog = wx.FileDialog(self, message="选择单个文件", wildcard=filesFilter, style=wx.FD_OPEN)
        dialogResult = fileDialog.ShowModal()
        if dialogResult != wx.ID_OK:
            return
        path = fileDialog.GetPath()
        self.__TextBox.SetLabel(path)
        global tempfilename, filename

        (filepath,tempfilename) = os.path.split(path)
        (filename,extention) = os.path.splitext(tempfilename)

        dir_path = 'E:\TestCode'+ '/' + tempfilename
        copyfile(path,dir_path)

    def OnClick(self, click):
        insert_data()
        self.SetLabel("ok")
        self.__TextBox.SetLabel("ok")

def insert_data():
    db = pymysql.connect(host="localhost", user="root",
                         passwd="123456",
                         db="work",
                         charset='utf8')
    book = xlrd.open_workbook(tempfilename)
    sheet = book.sheet_by_name("Sheet1")
    cursor = db.cursor()
    sql1 = "create table `%s`(工卡号 varchar (255) ,步骤  varchar (255)  primary key ,识别码 varchar (1024),插件 varchar (10000),程序内容 varchar (1024)) character set utf8" %(filename)

    cursor.execute(sql1)

    for i in range(1,sheet.nrows):
        nums = sheet.cell(i, 0).value  # 取第i行第0列
        steps = sheet.cell(i, 1).value
        sbm = sheet.cell(i, 2).value
        cj = sheet.cell(i, 3).value
        cxnr =sheet.cell(i, 4).value
        value = (nums, steps, sbm, cj, cxnr)
        print(value)
        sql = 'INSERT INTO  %s ' %filename + 'VALUES(%s,%s,%s,%s,%s)'
        cursor.execute(sql, value)  # 执行sql语句
        db.commit()
    cursor.close()  # 关闭连接
    db.close()