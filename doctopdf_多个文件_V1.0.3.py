# -*- coding: utf-8 -*-


print("{:=^60}".format ('Copyright'))           # 格式化输出：=号补齐空位，^居中显示，长度60，显示Copyright
print ("{}\n{}".format('AUTHOR: TINY YUAN', 'VERSION: V1.0.4'))       #格式化输出，{0}显示“XX”，换行，显示{1}"XX"
#补充了可视化按钮
print("{:=^60}".format ('Copyright'))

import os
import comtypes.client
import win32com.client
#import tkinter
from tkinter import *
from tkinter import Tk
from tkinter import filedialog
from tkinter import messagebox

def open_filefolder():
    #待补充信息
    pass


def getfilepath():
    try:
        filenames = filedialog.askopenfilenames(title='选择需要转换的文件')  #选取多个文件名
        return filenames             
    except Exception as error:
        print('获取路径失败', error)
        pass
def opendir():
    start_directory= filedir
    os.startfile(start_directory)

def close():        #关闭对话框
    root.destroy()
    os.system ('exit') 

def convert(in_file,in_filename,out_file):
    try:
        print('开始转换==>', in_file)
        if in_filename.endswith(('.doc','.docx')) :
            application ='Word.Application' # 保存格式
            office = comtypes.client.CreateObject(application)
            doc = office.Documents.Open(in_file)
            doc.SaveAs(out_file,FileFormat=17)
            doc.Close()        
        elif in_filename.endswith(('.ppt','.pptx')) :
            application ='Powerpoint.Application' 
            office = comtypes.client.CreateObject(application)
            office.Visible = 1
            ppt = office.Presentations.Open(in_file)
            ppt.SaveAs(out_file,FileFormat=32)
            ppt.Close()  
        elif in_filename.endswith(('.xls','xlsx')) :
            # excel转化为pdf
            #excel = comtypes.client.CreateObject("Excel.Application"),saveas 出错
            office=win32com.client.DispatchEx("Excel.Application")
            office.Visible = 0
            xls = office.Workbooks.Open(in_file)
            xls.ExportAsFixedFormat(0, out_file,1,0)
            xls.Close()
        else:
            pass
        office.Quit()
        print('完成转换','\n')
    
    except Exception as error:
        print(in_file ,'转换失败', error)
        try:
            office.Quit()
        except:
            pass
        pass   

def main():
    filenames= getfilepath()
    
    try:
        global filedir
        filedir= os.path.split(filenames[0])[0]        #提取文件所在文件夹路径
    except:
        print ('请重新选择文件')
        messagebox.showinfo('错误','请重新选择文件')
    for file in filenames: 
        fullpath =os.path.normpath(file)            #待优化，getfilepath中统一normpath
        if not file.endswith(('.doc','docx','.ppt','.pptx','.xls','xlsx')):
            print ('文件格式不符')
            messagebox.showinfo('错误','文件格式不符')
        else:
            in_file = fullpath
            in_filename=os.path.split(in_file)[1] 
            ext=os.path.splitext(fullpath)[1]  #提取扩展名
            out_file=fullpath.replace(ext, '.pdf')  #文件名，用于print显示
            out_dir,out_filename=os.path.split(out_file)        #获取路径和文件名
            #out_filename=out_file.split('\')[-1]
            if os.path.exists(out_file):       
                print (out_file,'已存在，跳过转换')
                message = out_filename + '已存在，跳过转换'
                messagebox.showinfo('提示',message)                
                continue
            else:
                pass
            convert(in_file,in_filename,out_file)
    print ('转换结束')
    messagebox.showinfo('提示','转换结束')

font1=("time news roman", 7,"italic" )         #label字体 font=( 字体, 字号, 粗体, 斜体, 下划线, 删除线)  "italic"
root=Tk()
root.title ('选择文件所在位置')   #窗口名称title
root.geometry('350x140')    #设置窗口尺寸
label1=Label (root,text = '提示：').place(x=0,y=40)  #.grid (row=4,sticky=W)   #建立标签Label place(x=0,y=40,width=400,height=30)
label2=Label (root,text = '1. 点击“选择文件”按钮，同时选择需要转换的多个word文件').place(x=0,y=60)
label3=Label (root,text = '2. 耐心等待“转换完成”提示').place(x=0,y=80)
label4=Label (root,text = 'AUTHOR: TINY YUAN, V1.0.4',font=font1).place(x=90,y=120)
button1=Button (root, text='选择文件',width = 8, command = main).place(x=40,y=8)#.grid (row=3, column=1) #sticky=W
button2=Button(root, text='打开文件夹',width = 10,command = opendir).place(x=135,y=8)#.grid(row=3,column=20)
button3=Button(root, text='关闭',width = 8,command = close).place(x=240,y=8)#.grid(row=3,column=20)
root.mainloop() 

