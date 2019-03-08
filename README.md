# ExcelToLua

1.使用了两个开源的库 见引用路径
2.excelToLua 要求格式表头最少4行
  line1:字段名 line2:注释 line3:分割符号 line4:字段类型
3.字段类型目前实现的种类,其中s结尾的表示数组，用分割符号配置分割服默认是';'
  int int64 float string bool ints int64s floats strings bools 

4.详情参加demo test.xlsx->test.lua  toolBat.bat
