# -*- coding:utf-8 -*-
import os
from cgai_excel.Handler import get_excel_data


path = r'C:\Temp\output\otask_module.xlsx'
data = get_excel_data(path,extract_dirpath=r'C:\Temp\output\EXT')
print(data)





