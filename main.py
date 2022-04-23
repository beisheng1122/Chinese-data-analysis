from tkinter import *
from tkinter import filedialog
from docx import Document
from DataAnalysis import DateAnalysis
import win32com.client as wc
import os

class Interface(): 
    def __init__(self, master=None):
        self.Initialization()


    def CheckData():
        global Date
        Date = text.get(1.0,END).split("\t")

        if Date == ['\n']:
            Tip = Label(window,text = "提示：请输入文本，并点击校验数据按钮",font = ("Times",15))
            Tip.pack(side = TOP,anchor = W,padx = 10,pady = 10)

        else:
            Tip = Label(window,text = "提示：文本校验成功!",font = ("Times",15))
            Tip.pack(side = TOP,anchor = W,padx = 10,pady = 10)

            Interface.WriteData(Date)

    def ReadData():
        Tk.destroy(window)
        DateFileGUI()

    def WriteData(Date):
        with open("data.main","w",encoding="utf-8") as f:
            for str in Date:
                f.write(str)
                f.close()

        Tip = Label(window,text = "提示：读取成功! 请点击统计分析按钮",font = ("Times",15))
        Tip.pack(side = TOP,anchor = W,padx = 10,pady = 10)
        
    def Initialization():
        global window,text,Date
        window = Tk()

        window.iconbitmap('icon.ico')

        width = 800
        heigh = 800

        screenwidth = window.winfo_screenwidth() # 获取屏幕宽度
        screenheight = window.winfo_screenheight() # 获取屏幕高度

        # 屏幕宽高除以二得中心点坐标
        window.geometry('%dx%d+%d+%d'%(width, heigh, (screenwidth-width)/2, (screenheight-heigh)/2)) # 参数说明: 窗口宽x窗口高+窗口位于屏幕x轴+窗口位于屏幕y轴

        window.title('基于大数据的文本数据统计分析')

        label = Label(window,text = "请输入待统计分析文本：",font = ("Times",20))
        label.pack(side = TOP,anchor = W,padx = 10,pady = 10)

        buttonPath = Button(window,text = "选择文件",font = ("Times",10),command = Interface.ReadData)
        buttonPath.pack(side = TOP,anchor = E,padx = 10,pady = 10)

        text = Text(window,width = 100)
        text.pack(side = TOP,anchor = W,padx = 10,pady = 10)

        button = Button(window,text = "校验数据",font = ("Times",15),command = Interface.CheckData)
        button.pack(side = TOP,anchor = W,padx = 10,pady = 10)

        labelPath = Label(window,text = "请输入分析结果保存路径(默认保存桌面)：",font = ("Times",20))
        labelPath.pack(side = TOP,anchor = W,padx = 10,pady = 10)

        entryPath = Entry(window,width = 100)
        entryPath.pack(side = TOP,anchor = W,padx = 10,pady = 10)

        labelCloudBg = Label(window,text = "请输入词云背景图片路径(不填使用默认背景)：",font = ("Times",20))
        labelCloudBg.pack(side = TOP,anchor = W,padx = 10,pady = 10)

        entryCloudBg = Entry(window,width = 100)
        entryCloudBg.pack(side = TOP,anchor = W,padx = 10,pady = 10)

        labelFontPath = Label(window,text = "请输入字体路径(不填使用默认字体)：",font = ("Times",20))
        labelFontPath.pack(side = TOP,anchor = W,padx = 10,pady = 10)

        entryFontPath = Entry(window,width = 100)
        entryFontPath.pack(side = TOP,anchor = W,padx = 10,pady = 10)

        buttonSave = Button(window,text = "开始统计分析",font = ("Times",15),command = lambda:DateAnalysis.GetDate(0,entryCloudBg.get(),entryPath.get(),entryFontPath.get()))
        buttonSave.pack(side = TOP,anchor = W,padx = 10,pady = 10)

class DateFileGUI():
    def __init__(self, master = None):
        global Subwindow

        Subwindow = Tk()

        Subwindow.iconbitmap('icon.ico')

        width = 800
        heigh = 800

        screenwidth = Subwindow.winfo_screenwidth() # 获取屏幕宽度
        screenheight = Subwindow.winfo_screenheight() # 获取屏幕高度

        # 屏幕宽高除以二得中心点坐标
        Subwindow.geometry('%dx%d+%d+%d'%(width, heigh, (screenwidth-width)/2, (screenheight-heigh)/2)) # 参数说明: 窗口宽x窗口高+窗口位于屏幕x轴+窗口位于屏幕y轴

        Subwindow.title('基于大数据的文本数据统计分析')
            
        label = Label(Subwindow,text = "请输入待统计分析文本路径：",font = ("Times",20))
        label.pack(side = TOP,anchor = W,padx = 10,pady = 10)

        var = StringVar()

        entry = Entry(Subwindow,width = 100,textvariable = var)
        entry.pack(side = TOP,anchor = W,padx = 10,pady = 10)

        buttonPath = Button(Subwindow,text = "选择文件",font = ("Times",10),command = lambda:DateFileGUI.GetPath(var))
        buttonPath.pack(side = TOP,anchor = E,padx = 10,pady = 10)

        labelPath = Label(Subwindow,text = "请输入分析结果保存路径：",font = ("Times",20))
        labelPath.pack(side = TOP,anchor = W,padx = 10,pady = 10)

        entryPath = Entry(Subwindow,width = 100)
        entryPath.pack(side = TOP,anchor = W,padx = 10,pady = 10)

        labelCloudBg = Label(Subwindow,text = "请输入词云背景图片路径(不填使用默认背景)：",font = ("Times",20))
        labelCloudBg.pack(side = TOP,anchor = W,padx = 10,pady = 10)

        entryCloudBg = Entry(Subwindow,width = 100)
        entryCloudBg.pack(side = TOP,anchor = W,padx = 10,pady = 10)

        labelFontPath = Label(Subwindow,text = "请输入字体路径(不填使用默认字体)：",font = ("Times",20))
        labelFontPath.pack(side = TOP,anchor = W,padx = 10,pady = 10)

        entryFontPath = Entry(Subwindow,width = 100)
        entryFontPath.pack(side = TOP,anchor = W,padx = 10,pady = 10)

        buttonSave = Button(Subwindow,text = "开始统计分析",font = ("Times",15),command = lambda:DateAnalysis.GetDate(0,entryCloudBg.get(),entryPath.get(),entryFontPath.get()))
        buttonSave.pack(side = TOP,anchor = W,padx = 10,pady = 10)

        buttonBack = Button(Subwindow,text = "返回",font = ("Times",15),command = DateFileGUI.Back)
        buttonBack.pack(side = TOP,anchor = E,padx = 10,pady = 10)

        mainloop()

    def GetPath(var):
        Filepath = filedialog.askopenfilename() #获得选择好的文件
        var.set(Filepath)

        if DateFileGUI.CheckPathisDocx(Filepath):
            DateFileGUI.ReadDocx(Filepath)

        else:
            DateFileGUI.ReadFile(Filepath)
    
    def CheckPathisDocx(Filepath):
        if Filepath.endswith(".docx"):
            return True

        if Filepath.endswith(".doc"):
            newpath = DateFileGUI.DocSomethingDocx(Filepath)
            DateFileGUI.ReadDocx(newpath)

        else:
            return False
    
    def ReadDocx(Filepath):
        document = Document(Filepath)

        for p in document.paragraphs:
            with open("read.main","a",encoding="utf-8") as f:
                f.write(p.text)
                f.close()

        labelTip = Label(Subwindow,text = "提示：读取成功!",font = ("Times",15))
        labelTip.pack(side = TOP,anchor = W,padx = 10,pady = 10)

    def DocSomethingDocx(Filepath): # 将doc转换为docx
        Filepath = os.path.abspath(Filepath) # 获取文件绝对路径

        w = wc.Dispatch('Word.Application') # 创建word应用程序
        w.Visible = 0 # 后台运行
        w.DisplayAlerts = 0 # 不显示任何警告信息。如果为true那么在出现警告时它会停止运行。
        doc = w.Documents.Open(Filepath) # 打开文件
        newpath = os.path.splitext(Filepath)[0] + '.docx' # 新的文件路径
        doc.SaveAs(newpath, 12, False, "", True, "", False, False, False, False) # 保存为docx
        doc.Close()
        w.Quit()
        os.remove(Filepath) # 删除doc文件

        return newpath


    def ReadFile(Filepath):
        global date

        with open(Filepath,"r",encoding="utf-8") as f:
            date = f.read()
            f.close()

        with open("date.read","w",encoding="utf-8") as f:
            f.write(date)
            f.close()

        labelTip = Label(Subwindow,text = "提示：读取成功!",font = ("Times",15))
        labelTip.pack(side = TOP,anchor = W,padx = 10,pady = 10)

    def Back():
        Subwindow.destroy()
        Interface.Initialization()
        mainloop()

if __name__ == '__main__':
    Interface.Initialization()

    mainloop()
