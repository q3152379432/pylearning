#source:https://blog.csdn.net/BADAO_LIUMANG_QIZHI/article/details/89556589
from PyPDF2 import PdfFileReader, PdfFileWriter
 
 
def merge_pdfs(paths, output):
    pdf_writer = PdfFileWriter()
    for path in paths:      #逐个文件写入
        pdf_reader = PdfFileReader(path)
        for page in range(pdf_reader.getNumPages()):    #逐页写入 pdf_reader
        # 将每页添加到writer对象
            pdf_writer.addPage(pdf_reader.getPage(page))    
 
        #写入合并的pdf
    with open(output, 'wb') as out:
        pdf_writer.write(out)
 
if __name__ == '__main__':
    paths = [r'D:\Users\Yuan GA\desktop\测试文件夹\1.pdf', r'D:\Users\Yuan GA\desktop\测试文件夹\2.pdf']
    merge_pdfs(paths, output=r'D:\Users\Yuan GA\desktop\测试文件夹\merged.pdf')