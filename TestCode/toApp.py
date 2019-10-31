import wx
from trans import MainFrame
import sys


#主应用程序的核
class ShellApp(wx.App):

    def OnInit(self):
        mainFrame = MainFrame()
        mainFrame.Show(True)
        return True



if __name__ == '__main__':
    app = ShellApp()

    app.MainLoop()