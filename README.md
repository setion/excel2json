本代码在 https://github.com/paceaux/xlsx-to-json 的基础上进行更改

1. 增加对数据类型的处理，当前从第4行开始读取数据

典型数据格式如下：

   第一行：字段名
   
   第二行：字段类型

   第三行：字段名解释

![image](https://github.com/setion/excel2json/assets/3980802/2356a788-48da-42ae-95fb-ee031268298e)

   目前支持的类型有：number/int, float, array/list, boolean/bool, string

2. 支持整表导出和按sheet导出
3. 本次修改基于自身项目需求更改，若无法满足其他项目，还望自行修改。
   
# xlsx-to-json
Using Python, convert an Excel (xlsx or xls) document to JSON


## dependencies
* pylightxl
* xlrd

you can run `pip3 install pylightxl xlrd` in your terminal/command window if you don't have these dependencies installed

## Usage

* Run the command `python3 xlstojson.py` in your terminal/command window
* Do what the prompt says (which is enter the path to the file)

A .json file matching the name of your xlsx document will be generated.


## Caveats and Whatnots
I wrote this six years ago; it was the first thing I'd ever written in Python.  This code was written about as well as a non-Python programmer programming Python could write it.

I have made attempts to improve it, but I may still not know what I'm doing


