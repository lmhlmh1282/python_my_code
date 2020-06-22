import os
import io
import sys
import subprocess
import pandas as pd

#app_filepath:自己填写的打开应用  description:要打开链接在excel文件里的描述信息
app_filepath="D:\\chrome.lnk"   #用什么打开链接
excel_file_basename="课程链接.xlsx" #什么excel文件 
description="《概率论与数理统计》教学视频全集（宋浩）"   #要打开的链接对应的description
#excel文件夹
excel_filepath_dir="D:\\临时作业"
#excel地址
excel_filepath=os.path.join(excel_filepath_dir,excel_file_basename)

#筛选处理，并用命令行打开网页链接
def get_to_links(sheet):
    ################################################################################
    # apply筛选
    # def is_equal(x,str):
    #     return (x==str)
    # sheet=sheet.loc[sheet['A'].apply(is_equal,args=(description,))].reset_index(drop=True)
    #
    ######################################################################### 
    #筛选，reset_index是为了将索引从0开始。默认筛选后没有变化
    sheet=sheet.loc[(sheet["A"]==description)].reset_index(drop=True)
    # print(sheet)

    #判断dataframe是否为空
    is_empty=(len(sheet.index)==0)
    #print(is_empty)
    #########################################################################
    if(is_empty==True):
        #当前表没有该值，查询为空，啥都不做
        return
    # print(sheet)
    ##########################################################################
    #打开链接
    #运行命令,第0行第B列,注意索引要从0开始
    web_link=sheet["B"][0]
    main_str="{} {}".format(app_filepath,web_link)
    # print(main_str)
    ##################################################################################
    subprocess.Popen(main_str,shell=True)

#.def




#################################################################################
if __name__ == "__main__":
    #循环读取各个表格
    excel_file=pd.ExcelFile(excel_filepath)
    for sheet_name in excel_file.sheet_names:
        #pandas读取数据esheet_name
        sheet=pd.read_excel(excel_filepath,sheet_name=sheet_name,header=None)
        #添加表头 A B C ...     ,'A' is 65
        sheet.columns=[chr(65+i) for i in range(0,sheet.shape[1])]
        #进行处理
        get_to_links(sheet)
        #释放内存
        del sheet
    del excel_file