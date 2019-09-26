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
from tkinter import *
import tkinter.messagebox  # 要使用messagebox先要导入模块
from tkinter.scrolledtext import ScrolledText
from tkinter import INSERT
from tkinter import END
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
def process(client, appkey, token, text, audio_name):
    callback = MyCallback(audio_name)
    synthesizer = client.create_synthesizer(callback)
    synthesizer.set_appkey(appkey)
    synthesizer.set_token(token)
    synthesizer.set_voice('siyue')
    synthesizer.set_text(text)
    synthesizer.set_format(TTSFormat.WAV)
    synthesizer.set_sample_rate(TTSSampleRate.SAMPLE_RATE_8K)
    synthesizer.set_volume(50)
    synthesizer.set_speech_rate(0)
    synthesizer.set_pitch_rate(0)
    try:
        ret = synthesizer.start()
        if ret < 0:
            return ret
        synthesizer.wait_completed()
    except Exception as e:
        print(e)
    finally:
        synthesizer.close()
def process_multithread(client, appkey, token, number):
    thread_list = []
    for i in range(0, number):
        text = "这是线程" + str(i) + "的合成。"
        audio_name = "sy_audio_" + str(i) + ".wav"
        thread = threading.Thread(target=process, args=(client, appkey, token, text, audio_name))
        thread_list.append(thread)
        thread.start()
    for thread in thread_list:
        thread.join()
def toVoice(appkey,token,text,audio_name):
    client = ali_speech.NlsClient()
    # 设置输出日志信息的级别：DEBUG、INFO、WARNING、ERROR
    client.set_log_level('INFO')
    process(client, appkey, token, text, audio_name)
    # 多线程示例
    # process_multithread(client, appkey, token, 2)



def getConfig():
    return {
        "appkey": 'TXH74wgv56oGoFFF',
        "token": '97c5f8ab3f2a4b7c91804ec9605d5c3e'
    }

cf = getConfig()
appkey = cf['appkey']
token = cf['token']


def changeVoice(load_dict,voiceDir,file):
        dir = voiceDir+'/'+file+'/';
        if os.path.exists(dir):
            print('exists dir%s'%(dir))
        else: 
            os.makedirs(dir)                   
        for item in load_dict:   
             toVoice(appkey,token,item['text'],dir+item['name']+'.wav')




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
                        changeVoice(load_dict, path.get(), projectName.get())
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

initWindow()

