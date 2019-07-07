# WorkSupport概述
用于办公辅助的程序，将常用功能集成在GUI中。
整个框架分4部分
- 切换函数（用于各功能选择时的控件切换）
- 功能函数
- 菜单
- 功能控件
如需添加功能，需要分别4部门中添加对应内容。


# 主要功能
- excel操作
  - [x] 数据对比
  - [x] 合并Excel
  - [ ] 非标数据合并

- 数据格式转换
  - [ ] Excel转DBF4
  - [x] Excel转（",","）
  
- 系统操作
  - [ ] 定时关机
  - [ ] 定时关IU
 
- 微信操作
  - [ ] 天气预报
  - [x] 获取群成员名单
  - [x] 微信控制摄像头拍照

# 功能介绍

### 数据对比
1、将要对比的数据放到一个Excel工作的前两个sheet中；  
2、确保要对比的字段顺序一致；  
3、选择文件，输入要对比的列（如：ABCD），点击开始对比。  
待完成：
文件输出后询问是否直接打开文件

### 合并Excel
1、将要合并的多个Excel放在同一个目录下，确保文件名不以“alldata”开头；  
2、要合并的数据表在sheet1，表头行数要相等；  
3、选择目录，输入表头行数，点击开始合并，合并结果会输出到“alldata(n).xls”文件中。   
待完成：
文件输出后询问是否直接打开文件

### Excel转（",","）
sql子查询经常要用到 where 字段名 in （",","） 格式，  
该功能将Excel(目前只支持.xlsx格式)中的数据转换成 （",","） 格式输出。

1、选择.xlsx文件；  
2、按要求格式输入要获取的单元格范围；  
3、输出结果直接复制即可。

### 获取群成员名单
1、将聊天群保存到通讯录；  
2、扫码登陆微信；  
3、双击群名称。

扫码登陆后获取到的群名单为活跃群及保存到通讯录的群，  
因微信问题非活跃群显示人数往往不准确，但不影响具体名单的获取。  
名单可输出到Excel，输出之后提示是否直接打开文件。

### 微信控制摄像头拍照
1、扫码登陆微信；  
2、在文件传输助手中发送"cap"获取图像。
