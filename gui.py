from tkinter import *
import tkinter.filedialog
import tkinter.messagebox

from spider import Spider


class Application(Frame):

    def __init__(self, master=None):
        Frame.__init__(self, master, width = 500, height = 200)
        self.pack()
        self.filePath = StringVar()
        self.outputPath = StringVar()
        self.createWidgets()

    def createWidgets(self):

        # 关键词选择相关控件
        self.filePathLabel = Label(self, text='关键词文件：')
        self.filePathLabel.grid(row = 0, column = 0)

        self.filePathEntry = Entry(self, textvariable = self.filePath, width=50)
        self.filePathEntry.grid(row = 0, column = 1)

        self.selectButton = Button(self, text='选择', command=self.selectFile)
        self.selectButton.grid(row = 0, column = 2)

        # 输出路径选择相关控件
        self.outputfilePathLabel = Label(self, text='输出文件路径：')
        self.outputfilePathLabel.grid(row=1, column=0)

        self.outputfilePathEntry = Entry(self, textvariable = self.outputPath, width=50)
        self.outputfilePathEntry.grid(row=1, column=1)

        self.outputselectButton = Button(self, text='选择', command=self.selectPath)
        self.outputselectButton.grid(row=1, column=2)

        # 开始按钮
        self.startButton = Button(self, text='开始', command=self.start)
        self.startButton.grid(row = 2,  column = 1)

    def selectFile(self):
        filePath = tkinter.filedialog.askopenfilename()
        if filePath != '':
            self.filePath.set(filePath)

    def selectPath(self):
        outputPath = tkinter.filedialog.askdirectory()
        if outputPath != '':
            self.outputPath.set(outputPath)

    def start(self):

        if self.filePath.get() == '' or self.outputPath.get() == '':
            tkinter.messagebox._show("Error", message="please select a keyword file and output path!")
            return

        spider = Spider()
        spider.readKeyWord(filePath=self.filePath)
        spider.searchKeyWord()
        spider.createResultExcel(outputFilePath=self.outputPath)
        tkinter.messagebox._show("Success", message="已经完成了爬取！")

if __name__ == '__main__':
    app = Application()
    # 设置窗口标题:
    app.master.title('Spider')
    # 主消息循环:
    app.mainloop()