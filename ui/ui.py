# -*- coding: utf-8 -*-
import sys
sys.path.append(r'./package/')
import do
import json

from tkinter.filedialog import askdirectory
from tkinter import *
import tkinter.messagebox  # 要使用messagebox先要导入模块
from tkinter.scrolledtext import ScrolledText
from tkinter import INSERT
from tkinter import END


def setRTxt(self, txt):
    self.delete(0.0, END)
    self.insert(INSERT, txt+'\n')
    self.see(END)

def initWindow():
        window = Tk()
        width = 1024
        height = 768
        voiceDir = 'c:\\voice'
        window.title('CPMS话术转换小助手')
        window.geometry('%sx%s' % (width, height))
        def hit_me():
            txt = t.get("0.0", "end")
            if len(projectName.get()) == 0:
                tkinter.messagebox.showinfo(title='Hi', message='请输入话术名称')
            else:
                if len(path.get()) == 0:
                    tkinter.messagebox.showinfo(title='Hi', message='请选择语音保存路径')
                else: 
                    print('start...')
                    try:
                        load_dict = json.loads(txt)
                        jtxt = json.dumps(load_dict, indent=4, ensure_ascii=False)
                        setRTxt(t, jtxt)
                        do.changeVoice(load_dict, path.get(), projectName.get())
                        successTxt = '语音生成成功！语音文件保存路径--->%s\%s' % (
                            path.get(), projectName.get())
                        tkinter.messagebox.showinfo(title='Hi', message=successTxt)
                    except Exception as ex:
                        extxt = "您的数据格式有误！%s" % (ex)
                        tkinter.messagebox.showinfo(title='Hi', message=extxt)

        def selectPath():
            path_ = askdirectory()
            path.set(path_)
            voiceDir = path_
        frmUp = Frame()
        Label(frmUp, text="话术名称").grid(row=0, column=1)
        projectName = StringVar()
        Entry(frmUp, textvariable=projectName).grid(row=0, column=2)
        path = StringVar()
        path.set(voiceDir)
        Label(frmUp, text="语音路径:").grid(row=0, column=3)
        Entry(frmUp, textvariable=path).grid(row=0, column=4)
        Button(frmUp, text="路径选择", command=selectPath).grid(row=0, column=5)
        Button(frmUp, text='点我合成语音', bg='green', font=(
            'Arial', 14), command=hit_me).grid(row=0, column=6)
        frmUp.grid(row=0, column=0, padx=1, pady=3)
        frmDown = Frame()
        frmDown.grid(row=1, column=0, padx=1, pady=3)

        t = ScrolledText(frmDown, font=("隶书", 18))
        t.grid(row=2, column=0)
        jsonT = [
            {
                "name": "话术主流程",
                "text": ""
            },
            {
                "name": "主流程-接受",
                "text": ""
            },
            {
                "name": "主流程-拒绝（二次挽留）",
                "text": " "
            },
            {
                "name": "主流程-未听清",
                "text": ""
            },
            {
                "name": "邀约成功结束语",
                "text": ""
            },
            {
                "name": "邀约失败结束语",
                "text": ""
            }
        ]
        jsTxt = json.dumps(jsonT, indent=4, ensure_ascii=False)
        setRTxt(t, jsTxt)
        window.mainloop()
