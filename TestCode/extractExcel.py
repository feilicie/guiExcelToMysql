#-*-coding: UTF-8 -*-
import wx
import xlrd
import pymysql

# import importlib
# importlib.reload(sys) #出现呢reload错误使用
'''def open_excel():
    try:
        book = xlrd.open_workbook("HX.xlsx")  # 文件名，把文件与py文件放在同一目录下
    except:
        print("open excel file failed!")
    try:
        sheet = book.sheet_by_name("Sheet1")
        return sheet
    except:
        print("locate worksheet in excel failed!")
        # 连接数据库

try:
    db = pymysql.connect(host="localhost", user="root",
                         passwd="123456",
                         db="work",
                         charset='utf8')
except:
    print("could not connect to mysql server")'''

def insert_data():
    db = pymysql.connect(host="localhost", user="root",
                         passwd="123456",
                         db="work",
                         charset='utf8')

    book = xlrd.open_workbook("HX.xlsx")

    sheet = book.sheet_by_name("Sheet1")

    cursor = db.cursor()

    tablename = ''
    sql = "CREATE TABLE `" + tablename + "` (`工卡号` varchar(255) ,`步骤` varchar(255) ,`识别码` varchar(1024),`插件` varchar（10000）,`程序内容` varchar(1024)) ENGINE=MyISAM DEFAULT CHARSET=utf8"
    cursor.execute("DROP TABLE IF EXISTS `" + tablename + "`")
    cursor.execute(sql)


    for i in range(1,sheet.nrows):
        nums = sheet.cell(i, 0).value  # 取第i行第0列
        steps = sheet.cell(i, 1).value
        sbm = sheet.cell(i, 2).value
        cj = sheet.cell(i, 3).value
        cxnr =sheet.cell(i, 4).value
        value = (nums, steps, sbm, cj, cxnr)
        print(value)
        sql = "INSERT INTO hx VALUES(%s,%s,%s,%s,%s)"
        cursor.execute(sql, value)  # 执行sql语句
        db.commit()
    cursor.close()  # 关闭连接
    db.close()


class ButtonFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, -1, '导入数据', size=(200, 200))
        panel = wx.Panel(self, -1)
        self.button = wx.Button(panel, -1, "导入", pos=(50, 50))  # 构造按钮
        self.Bind(wx.EVT_BUTTON, self.OnClick, self.button)
        self.button.SetDefault()
        loadbutton = wx.Button(panel, -1, "选择文件夹", pos=(50, 20))  # 构造按钮
        loadbutton.Bind(wx.EVT_BUTTON, self.OnClick, self.button)


    def OnClick(self,click):
        insert_data()
        self.button.SetLabel("ok")

if __name__ == '__main__':
     app = wx.PySimpleApp()
     frame = ButtonFrame()
     frame.Show()
     app.MainLoop()

