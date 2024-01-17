import tkinter as tk  
from tkinter import END, Text, filedialog
import os
from docx2pdf import convert

class MyApp:  
    def __init__(self, root): 
        self.Button = tk.Button(root, text="选择文件",command=lambda:choice_file(self)) 
        self.Button.pack()  
        self.Button.place(x=20, y=45)  

        self.Button = tk.Button(root, text="清除文件",command=lambda:del_file(self))  
        self.Button.pack()  
        self.Button.place(x=20, y=120) 

        self.file_data_Text = Text(root,width=98, height=12)
        self.file_data_Text.place(x=90, y=30)

        self.Button = tk.Button(root, text="输出文件夹",command=lambda:res_files(self))  
        self.Button.pack()  
        self.Button.place(x=16, y=222) 

        self.fold_data_Text = Text(root,width=98, height=4)
        self.fold_data_Text.place(x=90, y=210)


        self.Button = tk.Button(root, text="开始转换",command=lambda:deal(self))  
        self.Button.pack()  
        self.Button.place(x=20, y=310) 

        self.log_data_Text = Text(root,width=108, height=10)
        self.log_data_Text.place(x=20, y=360)


def choice_file(self):
    global a
    file_paths = filedialog.askopenfilenames()  # 打开文件对话框，让用户选择多个文件  
    if file_paths:  # 如果用户选择了文件  
        for file_path in file_paths:  
            a.append(file_path)
            file_write(self,0,file_path)
            # log_write(self,file_path)
    file_write(self,0,'当前已读取到下列文件：')
    for i in a:
        b = i.split('.')[-1]
    # b = "aaa"
    # log_write(self,a)
        if (b != "docx") & (b != "doc"):
            continue
        file_write(self,0,i)
    
def del_file(self):
    global a
    a = []
    file_write(self,1,"已清空！")

def file_write(self,num,fw):
    if num == 1:
        self.file_data_Text.delete(1.0,END)
        self.file_data_Text.insert(END, f'{fw}\n')
    else:
        self.file_data_Text.insert(END, f'{fw}\n')

def word_to_pdf(self,word_file, pdf_file):  
    # 确保文件存在  
    if not os.path.exists(word_file):  
        log_write(self,f"文件 {word_file} 不存在!")
        return  
    convert(word_file, pdf_file)
    # 将Word文档转换为PDF 
    log_write(self,f"已将 {word_file} 转换为 {pdf_file}") 

def log_write(self,fw):
    # self.log_data_Text.delete(1.0,END)
    self.log_data_Text.insert(END, f'{fw}\n')

def deal(self):
    global a
    for i in a:
        b = i.split('.')[-1]
    # b = "aaa"
    # log_write(self,a)
        if (b != "docx") & (b != "doc"):
            continue
        
        log_write(self,"当前处理文件为：\n" + i)
        if fold == '':
            log_write(self,"请选择输出文件夹")
            break
        ii = fold + i.split('/')[-1]
        # print(ii)
        if os.path.exists(ii.replace(b,"pdf")):
            os.remove(ii.replace(b,"pdf"))
            log_write(self,ii + "的原文件已删除")
            
        word_to_pdf(self,i,ii.replace(b,"pdf"))
    log_write(self,"已全部完成")


def res_files(self):
    global fold
    selected_folder = filedialog.askdirectory() # 打开文件对话框，让用户选择个文件夹  
    if selected_folder:
        fold = selected_folder + '/'
        fold_write(self,fold)

def fold_write(self,fw):
    self.fold_data_Text.delete(1.0,END)
    self.fold_data_Text.insert(END, f'{fw}\n')



if __name__ == "__main__":  
    a = []
    fold = ''
    root = tk.Tk()  
    root.title("WORD转PDFv1.0   By YukeLing")
    root.geometry("800x600") 
    root.resizable(False, False)
    app = MyApp(root)  
    root.mainloop()