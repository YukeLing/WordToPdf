import os  
from docx2pdf import convert  
  
def word_to_pdf(word_file, pdf_file):  
    # 确保文件存在  
    if not os.path.exists(word_file):  
        print(f"文件 {word_file} 不存在!")  
        return  
      
    # 将Word文档转换为PDF  
    convert(word_file, pdf_file)  
    print(f"已将 {word_file} 转换为 {pdf_file}")  

# 使用示例  
## word_to_pdf("example.docx", "output.pdf")

# 获取当前文件夹路径  
current_folder_path = os.getcwd()  
a = []
# 遍历当前文件夹中的文件和子文件夹  
for item in os.listdir(current_folder_path):  
    # 获取文件或子文件夹的完整路径  
    item_path = os.path.join(current_folder_path, item)  
    # 判断是否为文件  
    if os.path.isfile(item_path):  
        # 打印文件名  
        # print(f"文件: {item}")  
        # print(item_path)
        if item_path.split('.')[-1] == "docx":
            a.append(item_path)
        else:
            continue
    # 判断是否为文件夹  
    elif os.path.isdir(item_path):  
        continue

for i in a:
    b = i.split('.')[-1]
    if os.path.exists(i.replace(b,"pdf")):
        os.remove(i.replace(b,"pdf"))
        print("原文件已删除")
        
    print("当前处理文件为：" + i)
    word_to_pdf(i,i.replace(b,"pdf"))
    
print('处理完成！！！')