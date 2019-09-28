# -*- coding: utf-8 -*-
from urllib.request import urlretrieve
from xlrd import open_workbook
from xlutils.filter import XLRDReader
import os
import json
from pathlib import Path
import threading
import ali_speech
from ali_speech.callbacks import SpeechSynthesizerCallback
from ali_speech.constant import TTSFormat
from ali_speech.constant import TTSSampleRate
from docx import Document
from aliyunsdkcore.client import AcsClient
from aliyunsdkcore.request import CommonRequest
import tkinter as tk
from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilename
from tkinter import *
from tkinter import ttk
import tkinter.messagebox  # 要使用messagebox先要导入模块
import time
wordPath = 'C:\\Users\\ppx\\Desktop\\testword'
appkey = 'tQzL0nTf2p73E1Mi'
akID = 'LTAI4Fgat3CA8k39ofbrEuoW'
akSecret = 'gKBnGKPPvI6WWV4KJkAEFCbGuTPTV1'
savePath = 'c:\\voice'
# aixia,xiaowei,sicheng
person = 'aixia'


def getToken():
        # 创建AcsClient实例
    client = AcsClient(
        akID,
        akSecret,
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
    synthesizer.set_volume(90)
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
        thread = threading.Thread(target=process, args=(
            client, appkey, token, text, audio_name, voice))
        thread_list.append(thread)
        thread.start()
    for thread in thread_list:
        thread.join()


def toVoice(appkey, token, text, audio_name, voice):
    client = ali_speech.NlsClient()
    # 设置输出日志信息的级别：DEBUG、INFO、WARNING、ERROR
    client.set_log_level('INFO')
    process_multithread(client, appkey, token, text, audio_name, voice)


def changeVoice(load_dict, voiceDir, file, voice):
    dir = voiceDir+'/'+file+'/'
    if os.path.exists(dir):
        print('exists dir%s' % (dir))
    else:
        os.makedirs(dir)
    for item in load_dict:
        print(item['name'])
        print(item['text'])
        toVoice(appkey, getToken(), item['text'],
                dir+item['name']+'.wav', voice)


def textToVoice(load_dict, savePath, voiceName, person):
    print('start...')
    try:
        changeVoice(load_dict, savePath,
                    voiceName, person)
        successTxt = '语音生成成功！语音文件保存路径--->%s\%s' % (
            savePath, voiceName)
        print(successTxt)
    except Exception as ex:
        extxt = "您的数据格式有误！%s" % (ex)
        print(extxt)
# 解析word中的表格


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
# 解析word


def parseWord(file):
    f = open(file, 'rb')
    document = Document(f)
    list = []
    for t in document.tables:
        lt = handleTable(t)
        list = list+lt
    f.close()
    voiceName = os.path.basename(file)
    voiceName = voiceName.replace('.docx', '')
    textToVoice(list, savePath, voiceName, person)


def cbk(a, b, c):
    '''''回调函数 
    @a:已经下载的数据块 
    @b:数据块的大小 
    @c:远程文件的大小 
    '''
    per = 100.0*a*b/c
    if per >= 100:
        per = 100
    print('%.2f%%' % per)


def downloadfile(url, name):
    dir = os.path.abspath(wordPath)
    work_path = os.path.join(dir, name)
    filename, info = urlretrieve(url, work_path, cbk)
    parseWord(filename)


def parseExcel(file):
    workbook = open_workbook(file)
    sheet_names = workbook.sheet_names()
    sheets_object = workbook.sheets()
    sheet1_object = workbook.sheet_by_index(0)
    nrows = sheet1_object.nrows
    ncols = sheet1_object.ncols
    for r in range(1, nrows):
        id = sheet1_object.cell_value(rowx=r, colx=0)
        project_name = sheet1_object.cell_value(rowx=r, colx=1)
        order_name = sheet1_object.cell_value(rowx=r, colx=2)
        file_url = sheet1_object.cell_value(rowx=r, colx=3)
        # task_id = sheet1_object.cell_value(rowx=r, colx=0)
        # project_id = sheet1_object.cell_value(rowx=r, colx=0)
        print('%s,%s,%s' % (project_name, order_name, file_url))
        downloadfile(file_url, '%s-%s.docx' % (project_name, order_name))



def hitMe():
    if len(listPath.get()) == 0:
        tkinter.messagebox.showinfo(title='Hi', message='请输入话术列表文件')
    else:
        if len(wordSavePath.get()) == 0:
            tkinter.messagebox.showinfo(title='Hi', message='请选择话术word所在路径')
        else:
            if len(voiceSavePath.get()) == 0:
                tkinter.messagebox.showinfo(title='Hi', message='请选择语音保存路径')
            else:
                parseExcel(listPath.get());


def selectListFile():
    path_ = askopenfilename()
    listPath.set(path_)


def selectWord():
    path_ = askdirectory()
    wordSavePath.set(path_)
    wordPath = path_;


def selectVoicePath():
    path_ = askdirectory()
    voiceSavePath.set(path_)
    savePath = path_


root = tk.Tk()
width = 1024
height = 200
root.title('CPMS话术转换小助手')
root.geometry('%sx%s' % (width, height))
frmUp = Frame()
listPath = StringVar()
Label(frmUp, text="话术列表:").grid(row=0, column=1)
Entry(frmUp, textvariable=listPath).grid(row=0, column=2)
Button(frmUp, text="选择话术excel", command=selectListFile).grid(row=0, column=3)
wordSavePath = StringVar()
wordSavePath.set(wordPath)
Label(frmUp, text="word路径:").grid(row=0, column=4)
Entry(frmUp, textvariable=wordSavePath).grid(row=0, column=5)
Button(frmUp, text="选择word", command=selectWord).grid(row=0, column=6)
voiceSavePath = StringVar()
voiceSavePath.set(savePath)
Label(frmUp, text="语音路径:").grid(row=0, column=7)
Entry(frmUp, textvariable=voiceSavePath).grid(row=0, column=8)
Button(frmUp, text="路径选择", command=selectVoicePath).grid(row=0, column=9)


frmUp.grid(row=0, column=0, padx=1, pady=3)
frmDown = Frame()
Label(frmDown, text="选择声优:").grid(row=0, column=1)
comvalue = StringVar()  
comboxlist = ttk.Combobox(frmDown, textvariable=comvalue) 
comboxlist["values"] = ("aixia", "sicheng")
comboxlist.current(0)  # 选择第一个
comboxlist.grid(row=0, column=2)
Button(frmDown, text='点我合成语音', bg='green', font=(
    'Arial', 14), command=hitMe).grid(row=0, column=3)
frmDown.grid(row=1, column=0, padx=1, pady=3)
frmProcess = Frame()

# 设置下载进度条
tk.Label(frmProcess, text='下载进度:', ).grid(row=0, column=1)
canvas = tk.Canvas(frmProcess, width=465, height=22, bg="white")
frmDown.grid(row=2, column=0, padx=1, pady=3)

root.mainloop()



