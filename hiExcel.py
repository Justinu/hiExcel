#coding = utf-8

from tkinter import Tk, filedialog, messagebox
from appendExcel import appendNewExcel

root = Tk()
root.title("hiExcel")

while True:

    filelist = []
    
    #提示用户选择两个文件以上，尽量避免错误
    print('温馨提醒：至少选择2个以上的文件，才能进行合并操作！')

    #弹出对话框提示用户选择目标Excel文档，并添加进filelist列表中
    files = filedialog.askopenfilename(title = '选择目标Excel文档', filetypes = [('Excel文件2007', '.xlsx'), ('Excel文件2003', '.xls')])

    #如果完成了选择目标Excel文档步骤，则进入选择需合并至目标Excel的文档步骤，直至点击“取消”结束该循环
    while files != '':
        print(('您已选择Excel文档"{0}"').format(files))
        filelist.append(files)
        files = filedialog.askopenfilename(title = '选择需要合并至目标Excel的文档,如完成选择请点击“取消”',  filetypes = [('Excel文件2007', '.xlsx'), ('Excel文件2003', '.xls')])

    #检查完成文档选择后filelist列表中的元素个数，如果小于2，则提示用户“是否需要退出程序”,需要退出，则终止循环，程序结束；不需要，则继续循环
    if len(filelist) < 2:
        if messagebox.askyesno(title = '点击“确定”退出程序', message = '需要退出程序吗？'):
            break
        
    #检查完成文档选择后filelist列表中的元素个数，如果大于等于2，则进行Excel的合并操作，完成合并后提示合并完成
    elif len(filelist) >= 2:
        print('努力的进行文档合并中...')
        for index in range(1, len(filelist)):
            appendNewExcel(filelist[0], filelist[index])
            print('已完成"{0}"与"{1}"的合并...\n'.format(filelist[0], filelist[index]))
        print('文件合并完成!\n合并后的文件路径为"{0}"'.format(filelist[0]))

        #文档合并完成后询问用户是否继续有合并的需求，有需求，循环继续，无需求则终止循环，程序结束
        if not messagebox.askyesno(title = '继续吗？', message = '文档合并已完成，需要继续合并其他的Excel文件吗？'):
            break

root.destroy()
