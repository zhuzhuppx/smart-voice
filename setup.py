#!/usr/bin/python
"""
 * Copyright 2015 Alibaba Group Holding Limited
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
"""

try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

config = {
    'name': 'smart-voice',
    'version': '0.0.1',
    'description': 'smart-voice',
    'author': 'siyuan.zhu',
    'author_email': 'zhusiyuanhao@163.com',
    'license': 'Apache License 2.0',
    'url': 'https://github.com/zhuzhuppx/smart-voice.git',
    'install_requires': ['websocket-client', 'requests','python-docx','aliyun-python-sdk-core-v3','xlutils'],
    'packages': ['smart_voice'],
    'classifiers': (
        'Development Status :: 5 - Production/Stable',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: Apache Software License',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Topic :: Software Development',
    )
}

setup(**config)
