# cgai-excel

#### 介绍
操作excel

#### 图片获取  

规定： 
1. 图片必须以内嵌的形式插入到单元格当中  
2. 图片格子中的文字将不会识别  


#### 安装教程

```python
pip install cgai-excel

```

#### 使用说明
仅支持python3

```python
from cgai_excel.Handler import get_excel_data

path = r'F:\Temp\Q\excel_module.xlsx'
data = get_excel_data(path,extract_dirpath='F:\Temp\Q\Atemp')
```
