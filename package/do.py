# -*- coding: utf-8 -*-
import config
import tts
import json
import os
from pathlib import Path
cf = config.getConfig()
appkey = cf['appkey']
token = cf['token']


# def doChange(dir, file):
#     with open(dir+'/'+file, "r", encoding='utf-8-sig') as load_f:
#         load_dict = json.load(load_f)
#         for item in load_dict:
#             dir = 'voice/'+file.replace('.json', '')+'/';
#             my_file = Path(dir)
#             try:
#                 my_abs_path = my_file.resolve()
#             except FileNotFoundError:
#                 # 不存在
#                 os.makedirs(dir)
#             else:
#                 # 存在
#                 print('exists')
#             tts.toVoice(appkey, token, item['text'], dir+item['name']+'.wav')


def changeVoice(load_dict,voiceDir,file):
        dir = voiceDir+'/'+file+'/';
        if os.path.exists(dir):
            print('exists dir%s'%(dir))
        else: 
            os.makedirs(dir)                   
        for item in load_dict:   
             tts.toVoice(appkey,token,item['text'],dir+item['name']+'.wav')
