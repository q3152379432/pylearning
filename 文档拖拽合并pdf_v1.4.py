# -*- coding: utf-8 -*-
import os
import platform
import time
import comtypes.client
import win32com.client
from TkinterDnD2 import *
import tkinter.messagebox as Messagebox
from PyPDF2 import PdfFileReader, PdfFileWriter
import winreg
import fitz #用于img2pdf

try:
    from Tkinter import *
    from ScrolledText import ScrolledText
except ImportError:
    from tkinter import *
    from tkinter.scrolledtext import ScrolledText

def get_desktop():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,r'Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders')
    return winreg.QueryValueEx(key, "Desktop")[0]
 
def moveup():
    if len(listbox.curselection()) and listbox.curselection()[0]>0:
        pos = listbox.curselection()[0]
        text = listbox.get(pos)
        listbox.delete(pos)
        listbox.insert(pos-1,text)
        listbox.selection_clear("end")
        listbox.selection_set(pos-1)
 
def movedown():
    if len(listbox.curselection()) and listbox.curselection()[0]<listbox.size()-1:
        pos = listbox.curselection()[0]
        text = listbox.get(pos)
        listbox.delete(pos)
        listbox.insert(pos+1,text)
        listbox.selection_clear("end")
        listbox.selection_set(pos+1)
def update_currentstatus(current_statusinfo):
    current_status.set(current_statusinfo)
    root.update()

def convert():    
    pdf_path_list = []
    global waitfordelete
    waitfordelete = []
    for pathnum in range(listbox.size()):
        path = os.path.normpath(listbox.get(pathnum))
        inputfile_dir,inputfile_name = os.path.split(path)
        #print (path)
        current_statusinfo=inputfile_name+'转换中...'        
        update_currentstatus(current_statusinfo)
        try:
            if path.endswith((".pdf",".PDF")) :      #or str.endswith(path, ".PDF"):
                pdf_path_list.append(path)
            elif path.endswith((".doc",".docx",'DOC','DOCX','wps','WPS')):      #str.endswith(path,".imgdoc") or str.endswith(path, ".docx") or str.endswith(path, ".wps") or str.endswith(path,".DOC") or str.endswith(path, ".DOCX") or str.endswith(path, ".WPS"):
                office = win32com.client.DispatchEx("Word.Application")
                #office = comtypes.client.CreateObject("Word.Application")
                office.Visible = 0
                word = office.Documents.Open(path)
                out_file=path+".pdf"
                word.SaveAs(out_file, FileFormat=17)
                word.Close()
                #office.quit()
                office.Quit()
                pdf_path_list.append(out_file)
                waitfordelete.append(out_file)
            elif path.endswith((".ppt",".pptx",".PPT",".PPTX")):    #str.endswith(path,".ppt") or str.endswith(path, ".pptx") or str.endswith(path,".PPT") or str.endswith(path, ".PPTX"):
                office = win32com.client.DispatchEx("Powerpoint.Application")     #office = comtypes.client.CreateObject("Powerpoint.Application")
                office.Visible = 1 #ppt在转换的时候必须可视
                ppt = office.Presentations.Open(path)
                out_file=path+'.pdf'  #文件名，用于print显示
                print (out_file)
                ppt.SaveAs(out_file, FileFormat=32)
                ppt.Close()
                office.Quit()
                pdf_path_list.append(out_file)
                waitfordelete.append(out_file)
            elif path.endswith(('.xls','.xlsx','.XLS','.XLSX')):        #str.endswith(path,".xls") or str.endswith(path, ".xlsx") or str.endswith(path,".XLS") or str.endswith(path, ".XLSX"):
                office=win32com.client.DispatchEx("Excel.Application")
                office.Visible = 0
                excel = office.Workbooks.Open(path)
                out_file = path+".pdf"
                print (out_file)
                excel.ExportAsFixedFormat(0, out_file,1,0)
                excel.Close()
                office.Quit()
                pdf_path_list.append(out_file)
                waitfordelete.append(out_file)
            elif path.endswith(('.jpg','.JPEG','.png','.PNG','.bmp','BMP','tif','TIF','gif','GIF')):
                out_file=img2pdf(path)        #返回生成的pdf的全路径
                pdf_path_list.append(out_file)
                waitfordelete.append(out_file)
            else:
                print('不支持文件格式：',inputfile_name)
                current_statusinfo='不支持的文件格式：'+inputfile_name        
                update_currentstatus(current_statusinfo)
        except Exception as err:
            errinfo=inputfile_name+'转换出错：',err
            Messagebox.showinfo('错误',errinfo)

    output_filename = en1.get()     #手动输入的文件名
    if len(pdf_path_list):
        if not str.endswith(output_filename,".pdf") and not str.endswith(output_filename,".PDF"):   #添加扩展名
            output_fullname=output_filename+".pdf"
        else:
            output_fullname=output_filename

        global output_dir
        output_dir=os.path.split(path)[0]       #合并文件保存文件夹设置为最后一个文件所在位置
        global output_fullpath
        output_fullpath=os.path.join(output_dir,output_fullname)
        
        pdf_writer=merge_pdfs(pdf_path_list)        #读入文件
        
        with open(output_fullpath, 'wb') as output:
            pdf_writer.write(output)
            output.close()
        update_currentstatus(current_statusinfo)

        try:
            checkout=CheckVar1.get()    #获取复选框值
            if checkout==0:
                current_statusinfo='单个pdf文件删除中...'
                update_currentstatus(current_statusinfo)                
                delete_singlepdf()
            else:
                pass
        except:
            current_statusinfo='单个pdf文件删除失败'
            update_currentstatus(current_statusinfo)

        print('pdf合成结束')
        current_statusinfo=output_fullname+'  合成结束！'
        update_currentstatus(current_statusinfo)
        Messagebox.showinfo('通知','合成结束！')

    else:
        Messagebox.showinfo('失败通知','没有可以合并的PDF文件！')
    
def merge_pdfs(pdf_path_list):
    print('开始合成pdf')
    
    pdf_writer = PdfFileWriter()
    for path in pdf_path_list:      #逐个文件写入
        try:
            current_dir,current_name=os.path.split(path)
            current_statusinfo= current_name+'合入中'        
            update_currentstatus(current_statusinfo)
            #current_status.set(path_dir_name+' 合入中')
            pdf_reader = PdfFileReader(path)
            PageNumbers = pdf_reader.getNumPages()
            for page in range(PageNumbers):    #逐页写入 pdf_reader
                # 将每页添加到writer对象
                pdf_writer.addPage(pdf_reader.getPage(page))
            checkout2=CheckVar2.get()
            if (PageNumbers % 2)==1 and checkout2==1:
                pdf_writer.addBlankPage(width= None,height=None)    #为基数页时，添加空白页
            else:
                pass
        except Exception as err:
            errinfo=current_name+'合入失败：'+err
            Messagebox.showinfo()('错误',errinfo)

    return pdf_writer
    #写入合并的pdf
    #获取文本框里面的文件名
def img2pdf(path):  
    imgpdf=fitz.open()
    imgdoc = fitz.open(path)         # 打开图片
    pdfbytes = imgdoc.convertToPDF()    # 使用图片创建单页的 PDF
    img_pdfbytes = fitz.open("pdf", pdfbytes)
    imgpdf.insertPDF(img_pdfbytes)          # 将当前页插入文档
    out_file = path+".pdf"
    imgpdf.save(out_file)          # 保存pdf文件
    imgdoc.close()
    imgpdf.close()
    return out_file
    
def delete_singlepdf():     #删除过程中生成的单独pdf文件
    try:
        for pdf in waitfordelete:
            os.remove(pdf) 
        #Messagebox.showinfo('通知','单个pdf文件删除完成！')
    except:
        Messagebox.showinfo('错误','未找到生成的单个pdf文件')
        pass

def open_combinefolder():   #打开合成的pdf所在文件夹
    try:
        os.startfile(output_dir)
    except:
        Messagebox.showinfo('错误','未找到合成的pdf文件')
    
def open_combinefile():     #用默认程序打开合成pdf文件
    try:
        os.startfile(output_fullpath)
    except:
        Messagebox.showinfo('错误','未找到合成的pdf文件')
    
def deleteline():       #删除选中行
    listbox.delete(ACTIVE)

def clear_listbox():
    listbox.delete(0,END)


def test():
    checkout=CheckVar1.get()
    print(checkout)


root = TkinterDnD.Tk()
root.withdraw()
root.resizable(width=False, height=False)
root.title('PDF文件转换与合并工具 for yjzx')

current_status=StringVar()
current_status.set('待开始')
#global label2
 
root.grid_rowconfigure(1,  minsize=10)#(1, weight=1, minsize=350)
root.grid_columnconfigure(0, minsize=10)
 
#text_status=Text(root, width=50,height=2).grid(row=2, column=0, padx=10, pady=5,sticky=W)
labelframe1 = Frame(root)
labelframe1.grid(row=0, column=0, columnspan=1, pady=0,sticky=W)
label1_1=Label(labelframe1, text='请把需要合并的文件拖拽到下方框内，支持word/ppt/excel/pdf/图片。').pack(side=LEFT,padx=5)

labelframe1_2 = Frame(root)
labelframe1_2.grid(row=1, column=0, columnspan=1, pady=0,sticky=W)
#label2=Label(labelframe1, textvariable=current_status).pack(side=LEFT,padx=5)
label1_2=Label(labelframe1_2, text='当前运行状态：').pack(side=LEFT,padx=5)
label1_2=Label(labelframe1_2, textvariable=current_status).pack(side=LEFT,padx=0)

listbox = Listbox(root, name='dnd_demo_listbox',            #拖入框
                    selectmode='BROWSE', width=1, height=1)
listbox.grid(row=2, column=0, padx=10, pady=10, sticky='news')
#listbox.insert(END, os.path.abspath(__file__)) 
                   
buttonbox1 = Frame(root,width=20)
buttonbox1.grid(row=2, column=1, sticky=N,pady=5) #grid(row=2, column=1, columnspan=1, sticky=N,pady=5)
Button(buttonbox1, text='上移选中的行', width=10,command=moveup).pack(
                    side=TOP, padx=5,pady=5)
Button(buttonbox1, text='下移选中的行', width=10,command=movedown).pack(
                    side=TOP, padx=5,pady=5)
Button(buttonbox1, text='删除选中的行', width=10,command=deleteline).pack(
                    side=TOP, padx=5,pady=5)
Button(buttonbox1, text='清空列表', width=10,command=clear_listbox).pack(side=TOP, padx=5,pady=5)       #测试用                    
Button(buttonbox1, text='开始转换',width=10,command=convert).pack(side=TOP, padx=5,pady=5)
Button(buttonbox1, text='退出', width=10,command=root.quit).pack(side=TOP, padx=5,pady=5)


labelframe2 = Frame(root)
labelframe2.grid(row=3, column=0, columnspan=1, sticky=N,pady=5)#.grid(row=2, column=0, columnspan=1, pady=5)
label3 = Label(labelframe2,text="合并后文件名：",bd=0,height=1).pack(side=LEFT, padx=10)
en1 = Entry(labelframe2)
en1.pack(side=LEFT, padx=5)
en1.insert(0,"合成的pdf文件")
Button(labelframe2, text='打开合成的文件', command=open_combinefile,height=1).pack(side=LEFT, padx=5)
Button(labelframe2, text='打开所在位置', command=open_combinefolder,height=1).pack(side=LEFT, padx=5) 

CheckVar1=IntVar()
CheckVar2=IntVar()
checkframe3 = Frame(root)
checkframe3.grid(row=4, column=0, columnspan=1, sticky=N,pady=5)#.grid(row=2, column=0, columnspan=1, pady=5)
checkb1=Checkbutton(checkframe3, text = "删除生成的单个pdf文件", variable = CheckVar1, \
    onvalue = 0, offvalue = 1, height=1, width = 20).pack(side=LEFT, padx=5) 
checkb2=Checkbutton(checkframe3, text = "添加空白页至总页数为基数的pdf文件末尾", variable = CheckVar2, \
    onvalue = 1, offvalue = 0, height=1).pack(side=LEFT, padx=5) 
   
 
# make the Label a drop target:
def drop_enter(event):
    event.widget.focus_force()
    #print('Entering %s' % event.widget)
    return event.action
 
def drop_position(event):
    #print('Position: x %d, y %d' %(event.x_root, event.y_root))
    return event.action
 
def drop_leave(event):
    #print('Leaving %s' % event.widget)
    return event.action
 
def drop(event):
    if event.data:
        #print('Dropped data:\n', event.data)
        if event.widget == listbox:
            files = listbox.tk.splitlist(event.data)
            for f in files:
                if os.path.exists(f):
                    dr = True
                    for pathnum in range(listbox.size()):
                        if f == listbox.get(pathnum):
                            dr = False
                            break
                    if dr:
                        #print('Dropped file: "%s"' % f)
                        listbox.insert('end', f)
                else:
                    print('Not dropping file "%s": file does not exist.' % f)
        elif event.widget == text:
            # calculate the mouse pointer's text index
            bd = text['bd'] + text['highlightthickness']
            x = event.x_root - text.winfo_rootx() - bd
            y = event.y_root - text.winfo_rooty() - bd
            index = text.index('@%d,%d' % (x,y))
            text.insert(index, event.data)
        else:
            print('Error: reported event.widget not known')
    return event.action
 
listbox.drop_target_register(DND_FILES, DND_TEXT)
listbox.dnd_bind('<<DropEnter>>', drop_enter)
listbox.dnd_bind('<<DropPosition>>', drop_position)
listbox.dnd_bind('<<DropLeave>>', drop_leave)
listbox.dnd_bind('<<Drop>>', drop)
 
# make the Label a drag source:
 
def drag_init(event):
    #data = listbox['text']
    #return (COPY, DND_TEXT, data)
    pass
 
def drag_end(event):
    pass
 
listbox.drag_source_register(DND_TEXT)
listbox.dnd_bind('<<DragInitCmd>>', drag_init)
listbox.dnd_bind('<<DragEndCmd>>', drag_end)
 
root.update()
sw = root.winfo_screenwidth() #tkinter自带的获取屏幕宽度
sh = root.winfo_screenheight()#获取屏幕高度
ww = root.winfo_width()#获取程序窗口宽度
wh = root.winfo_height()#获取程序窗口高度，注意，在调用获取程序窗口宽高之前要先刷新，即调用root.update()，否则得到的是初始值，即0,0
 
x = (sw-ww)      #窗口位置
y = (sh-wh)/10
root.geometry("%dx%d+%d+%d" %(ww,wh,x,y))
#root.iconbitmap(os.getcwd()+"\\xc.ico")
root.update_idletasks()
root.deiconify()
root.mainloop()


'''
buttonbox2 = Frame(root)
buttonbox2.grid(row=3, column=0, columnspan=1, pady=0,sticky=W)
Button(buttonbox2, text='开始转换',width=10,command=convert).pack(side=LEFT, padx=5)
#Button(buttonbox2, text='删除生成的单个pdf文件', command=delete_singlepdf).pack(side=LEFT, padx=5) 
Button(buttonbox2, text='退出', width=10,command=root.quit).pack(side=LEFT, padx=5)
'''
