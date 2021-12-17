# -*- coding:utf-8 -*-
import os
from cgai_excel.Handler import get_excel_data


path = r'F:\Temp\Q\excel_module.xlsx'
data = get_excel_data(path,extract_dirpath='F:\Temp\Q\Atemp')
print(data)





