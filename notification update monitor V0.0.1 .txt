#!/usr/bin/env python
# -*- coding: utf-8 -*-

import requests
import re
import os
import time
import sys
from tkinter import messagebox
from tkinter import *
import smtplib
from email.mime.text import MIMEText
from email.header import Header

def sendmail(update):     #发送email
    host="smtp.qq.com"  #设置服务器
    username="315237943@qq.com"    #用户名
    password="piybuwjxiapwcbeb"   #授权码   piybuwjxiapwcbeb

    sender = '315237943@qq.com'     #发送邮件地址
    receiver = 'yuanga@huanke.com.cn' # 接收邮件，可设置为你的QQ邮箱或者其他邮箱
    
    subject = update           #邮件主题
    message= MIMEText ('结果查询链接：http://yz.tongji.edu.cn/zsxw.htm', 'plain', 'utf-8')
    message['Subject'] = Header (subject, 'utf-8')
    message ['From'] = '315237943@qq.com'     #发件人名称，自定义
    message ['To'] =  'yuanguoan@163.com'   #收件人名称，自定义

    try:
        server=smtplib.SMTP_SSL('smtp.qq.com', 465, timeout=5)      #SMTP默认端口为25；QQ邮箱465或587
        server.login(username, password)  
        server.sendmail(sender, receiver, message.as_string())
        print ("邮件发送成功")
    except Exception as error:
        print ("邮件发送失败", error)
    finally:
        server.quit()

def sendmailnoupdate(update):         #无更新发送email
    host="smtp.qq.com"  #设置服务器
    username="315237943@qq.com"    #用户名
    password="piybuwjxiapwcbeb"   #授权码   piybuwjxiapwcbeb

    sender = '315237943@qq.com'     #发送邮件地址
    receiver = 'yuanga@huanke.com.cn' # 接收邮件，可设置为你的QQ邮箱或者其他邮箱
    
    subject = '无更新'            #邮件主题
    message= MIMEText ('未检测到更新', 'plain', 'utf-8')
    message['Subject'] = Header (subject, 'utf-8')
    message ['From'] = '315237943@qq.com'     #发件人名称，自定义
    message ['To'] =  'yuanguoan@163.com'   #收件人名称，自定义

    try:
        server=smtplib.SMTP_SSL('smtp.qq.com', 465, timeout=5)      #SMTP默认端口为25；QQ邮箱465或587
        server.login(username, password)  
        server.sendmail(sender, receiver, message.as_string())
        print ("邮件发送成功")
    except Exception as error:
        print ("邮件发送失败", error)
    finally:
        server.quit()
    
'''
    try:
        smtpObj = smtplib.SMTP()
        smtpObj.connect(mail_host, 25)    # 25 为 SMTP 端口号
        smtpObj.login(username, password)  
        smtpObj.sendmail(sender, receivers, message.as_string())
        print ("邮件发送成功")
    except smtplib.SMTPException:
        print  ("Error: 无法发送邮件")
'''

def getHTMLText():       #获取相应url的HTML代码
    kv={'user-agent':'Mozilla/5.0'}     #模拟浏览器访问 Mozilla/5.0
    url = 'http://yz.tongji.edu.cn/zsxw.htm'
    try:
        r = requests.get(url, headers=kv, timeout = 30) #requests库get函数
        r.raise_for_status()        #出错退出
        r.encoding = r.apparent_encoding    #根据HTML内容确定编码格式
        return r.text       #r.text属性，HTML文本
    except:
        print ('网站读取失败')
        return ""

def getinfo(html):
    #机械与能源工程学院</span></td>   <td style="border: 0px rgb(0, 0, 0);background-color: transparent">​</td>
    #pat = r'机械与能源工程学院</span></td>.*?</td>'
    keyinfo= re.findall(r'2020年.*博士研究生.*复试', html)       #匹配，提取（）内关键信息
    #print (keyinfo)
    return keyinfo[0]

def showmessagebox(info):
    root=Tk()
    root.withdraw()
    result=messagebox.showinfo('提示','有更新'+ info)
    
  
def main ():
    i=1
    k=0
    html= getHTMLText()
    keyinfo1= getinfo(html)
    update=''
    while i<=432000:             #持续24小时
        html= getHTMLText()
        #print (html [:200])
        keyinfo2= getinfo(html)
        #result = re.search('月', keyinfo)
        if keyinfo2==keyinfo1:      #no update
            k=k+1
            i= i+1800
            print(time.ctime())
            print ('查看第', k , '次', ',暂无更新')
            #sendmailnoupdate(update)             #没有更新也需要提醒时使用
            time.sleep (3600)       #休眠60分钟
        else:       #got update
            #sendemail()
            #showmessagebox()
            update=keyinfo[0]
            sendmail(update)      #发送邮件
            #showmessagebox(keyinfo)            #需要弹框提醒时使用
            print ('有更新')
            keyinfo1=keyinfo2           #更新keynfo1的值
            time.sleep (3600)       #休眠60分钟
            #break
main()
    


    


