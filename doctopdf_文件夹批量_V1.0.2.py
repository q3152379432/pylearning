import os
import comtypes.client
#import tkinter
from tkinter import Tk
from tkinter import filedialog

print("{:=^60}".format ('Copyright'))           # 格式化输出：=号补齐空位，^居中显示，长度60，显示Copyright
print ("{}\n{}".format('AUTHOR: TINY YUAN', 'VERSION: V1.0.2'))       #格式化输出，{0}显示“XX”，换行，显示{1}"XX"
print("{:=^60}".format ('Copyright'))

def getfilepath():
    root=Tk()
    root.title ('选择文件所在位置')   #窗口名称title
    root.geometry('300x0')    #设置窗口尺寸
    try:
        filepath = filedialog.askdirectory(title='选择文件所在文件夹')
        #path=filepath
        filepath=os.path.normpath( filepath) #格式化路径，/ 转化为 \
        return filepath             
    except Exception as error:
        print('获取路径失败', error)
        pass
    finally:
        root.destroy()
        pass

def main():
    folder = getfilepath()
    out_folder=os.path.join(folder, r'pdf')        #保存pdf的文件位置
    if not os.path.exists(out_folder):
        os.mkdir(out_folder)    #创建pdf保存的文件夹
        print('创建pdf保存文件夹',out_folder)
    else: 
        pass
    #folder=r'D:\users\Yuan GA\document\python test'
    #for dirpath, dirnames, filenames in os.walk(folder):   遍历子文件夹，当需要将子目录文件转换时，嵌套循环
    filenames= os.listdir(folder)
    for file in filenames:
        if file.endswith('.doc') or file.endswith('docx'):
            fullpath = os.path.join(folder, file)
            wdFormatPDF = 17 # 保存格式
            in_file = fullpath
            ext=os.path.splitext(fullpath)[1]  
            out_file_name=file.replace(ext, '.pdf')  #文件名，用于print显示
            out_file=os.path.join(out_folder,out_file_name)
            if os.path.exists(out_file):       
                print (out_file_name,'已存在，跳过转换')
                continue
            else:
                pass

            try:
                print('开始转换==>', file)
                word = comtypes.client.CreateObject('Word.Application')
                doc = word.Documents.Open(in_file)
                doc.SaveAs(out_file, FileFormat=wdFormatPDF)
                doc.Close()
                word.Quit()
                print('完成转换==>', out_file_name, '\n')
                
            except Exception as error:
                doc.Close()
                word.Quit()
                print(file ,'转换失败', error)
                pass
        else:
            pass
    print ('转换结束')
    print('保存的位置为：',out_folder )

main()
os.system ('pause & exit')
