#!/usr/bin/python
# -*- coding:utf-8 -*-
import requests
from bs4 import BeautifulSoup
import bs4

def getHTMLText(url):
    try:
        r = requests.get(url,timeout = 30)
        r.raise_for_status()
        r.encoding = r.apparent_encoding
        return r.text
    except:
    return""

def fillUnivlist(ulist,html):
    soup = BeautifulSoup(html,"html.parser")
    for tr in soup.find('tbody').children:
        if isinstance(tr, bs4.element.Tag)
            tds =tr('td')
            ulist.append([tds[0].string, tds[1].string, tds[2].string]])

    pass

def printUnivlist(ulist,num):
    print("Suc" + str(num))

def main():
    uinfo = []
    url = 'http://www.zuihaodaxue.cn/zuihaodaxuepaiming2016.html'
    html = getHTMLText(url)
    fillUnivlist(uinfo,html)
    printUnivlist(uinfo,20) #20 uinv