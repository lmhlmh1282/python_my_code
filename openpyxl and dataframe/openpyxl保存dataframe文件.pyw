import sys
import subprocess
import pandas as pd
import os
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
#app_filepath:自己填写的打开应用  description:要打开链接在excel文件里的描述信息
excel_file_basename="课程链接.xlsx" #什么excel文件 
#excel文件夹
excel_filepath_dir="xxx\\课程连接"
#excel地址
excel_filepath=os.path.join(excel_filepath_dir,excel_file_basename)



#################################################################################
if __name__ == "__main__":
    excel_file=pd.ExcelFile(excel_filepath)
    wb = openpyxl.Workbook()
    for sheet_name in excel_file.sheet_names:
        ws=wb.create_sheet(sheet_name)
        #pandas读取数据esheet_name
        sheet=pd.read_excel(excel_filepath,sheet_name=sheet_name,header=None)
      
        #释放内存
        for r in dataframe_to_rows(sheet,index=True,header=True):
            ws.append(r)
        wb.save("测试.xlsx")
        