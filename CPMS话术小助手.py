# -*- coding: utf-8 -*-

import json
import os
from pathlib import Path
# -*- coding: utf-8 -*-
import threading
import ali_speech
from ali_speech.callbacks import SpeechSynthesizerCallback
from ali_speech.constant import TTSFormat
from ali_speech.constant import TTSSampleRate

from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilename
from tkinter import *
import tkinter.messagebox  # 要使用messagebox先要导入模块
from tkinter.scrolledtext import ScrolledText
from tkinter import INSERT
from tkinter import END
from tkinter import ttk

from docx import Document
# -*- coding: utf8 -*-
from aliyunsdkcore.client import AcsClient
from aliyunsdkcore.request import CommonRequest


def getToken():
        # 创建AcsClient实例
    client = AcsClient(
        "LTAI4Fgat3CA8k39ofbrEuoW",
        "gKBnGKPPvI6WWV4KJkAEFCbGuTPTV1",
        "cn-shanghai"
    )
    # 创建request，并设置参数
    request = CommonRequest()
    request.set_method('POST')
    request.set_domain('nls-meta.cn-shanghai.aliyuncs.com')
    request.set_version('2019-02-28')
    request.set_action_name('CreateToken')
    response = client.do_action_with_exception(request)
    token = json.loads(response)
    return token['Token']['Id']


def getConfig():
    return {
        "appkey": 'tQzL0nTf2p73E1Mi ',
        "token": getToken()
    }


def handleTable(t):
    jsonlist = []
    for i, row in enumerate(t.rows):
        row_content = []
        j = 0
        obj = {}
        for cell in row.cells:
            c = cell.text
            c = c.replace('\n', '')
            if j == 0:
                obj['name'] = c
            if j == 1:
                obj['text'] = c
            j = j+1
        jsonlist.append(obj)
    return jsonlist


def doParseWord(file):
    f = open(file, 'rb')
    document = Document(f)
    list = []
    for t in document.tables:
        lt = handleTable(t)
        list = list+lt
    f.close()
    return json.dumps(list, indent=4, ensure_ascii=False)


class MyCallback(SpeechSynthesizerCallback):
    # 参数name用于指定保存音频的文件
    def __init__(self, name):
        self._name = name
        self._fout = open(name, 'wb')

    def on_binary_data_received(self, raw):
        print('MyCallback.on_binary_data_received: %s' % len(raw))
        self._fout.write(raw)

    def on_completed(self, message):
        print('MyCallback.OnRecognitionCompleted: %s' % message)
        self._fout.close()

    def on_task_failed(self, message):
        print('MyCallback.OnRecognitionTaskFailed-task_id:%s, status_text:%s' % (
            message['header']['task_id'], message['header']['status_text']))
        self._fout.close()

    def on_channel_closed(self):
        print('MyCallback.OnRecognitionChannelClosed')


def process(client, appkey, token, text, audio_name, voice='aixia'):
    callback = MyCallback(audio_name)
    synthesizer = client.create_synthesizer(callback)
    synthesizer.set_appkey(appkey)
    synthesizer.set_token(token)
    # aixia,xiaowei,sicheng
    synthesizer.set_voice(voice)
    synthesizer.set_text(text)
    synthesizer.set_format(TTSFormat.WAV)
    synthesizer.set_sample_rate(TTSSampleRate.SAMPLE_RATE_16K)
    synthesizer.set_volume(60)
    synthesizer.set_speech_rate(-30)
    synthesizer.set_pitch_rate(-10)
    try:
        ret = synthesizer.start()
        if ret < 0:
            return ret
        synthesizer.wait_completed()
    except Exception as e:
        print(e)
    finally:
        synthesizer.close()

def process_multithread(client, appkey, token, text, audio_name, voice):
    thread_list = []
    for i in range(0, 5):     
        thread = threading.Thread(target=process, args=(client, appkey, token, text, audio_name, voice))
        thread_list.append(thread)
        thread.start()
    for thread in thread_list:
        thread.join()

def toVoice(appkey, token, text, audio_name, voice):
    client = ali_speech.NlsClient()
    # 设置输出日志信息的级别：DEBUG、INFO、WARNING、ERROR
    client.set_log_level('INFO')
    process_multithread(client, appkey, token, text, audio_name, voice)


cf = getConfig()
appkey = cf['appkey']
token = cf['token']


def changeVoice(load_dict, voiceDir, file, voice):
    dir = voiceDir+'/'+file+'/'
    if os.path.exists(dir):
        print('exists dir%s' % (dir))
    else:
        os.makedirs(dir)
    for item in load_dict:
        print(item['name'])
        print(item['text'])
        toVoice(appkey, token, item['text'], dir+item['name']+'.wav', voice)


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
                    changeVoice(load_dict, path.get(),
                                projectName.get(), comvalue.get())
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

    def selectWord():
        path_ = askopenfilename()
        wordPath.set(path_)
        jsTxt = doParseWord(path_)
        setRTxt(t, jsTxt)
    frmUp = Frame()
    wordPath = StringVar()
    Label(frmUp, text="word路径:").grid(row=0, column=1)
    Entry(frmUp, textvariable=wordPath).grid(row=0, column=2)
    Button(frmUp, text="选择word", command=selectWord).grid(row=0, column=3)
    Label(frmUp, text="话术名称").grid(row=0, column=4)
    projectName = StringVar()
    Entry(frmUp, textvariable=projectName).grid(row=0, column=5)
    path = StringVar()
    path.set(voiceDir)
    Label(frmUp, text="语音路径:").grid(row=0, column=6)
    Entry(frmUp, textvariable=path).grid(row=0, column=7)
    Button(frmUp, text="路径选择", command=selectPath).grid(row=0, column=8)
    comvalue = tkinter.StringVar()  # 窗体自带的文本，新建一个值
    comboxlist = ttk.Combobox(frmUp, textvariable=comvalue)  # 初始化
    comboxlist["values"] = ("aixia", "sicheng")
    comboxlist.current(0)  # 选择第一个
    comboxlist.grid(row=0, column=9)
    Button(frmUp, text='点我合成语音', bg='green', font=(
        'Arial', 14), command=hit_me).grid(row=0, column=10)
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


initWindow()
