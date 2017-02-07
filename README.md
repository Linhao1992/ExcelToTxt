# ExcelToTxt
  概述：将一个/多个客户的Excel转换成柜面批量业务的TXT

功能说明：
  1、自动识别Exce内的账户、姓名、金额信息，不用手工转格式
  2、支持批量转换，秒转多个批量文件
  3、账号位数错误会具体提示“某Excel的xx工作表的xx行错误”
  4、转换完成会自动提示"笔数、总金额"
  5、自动弹出生成的txt所在目录

Excel规范：已按照客户送来的格式规范进行最大限度兼容
  1、Excel要求：
      暂时仅支持xls格式
      每个Excel文件仅有一张批量信息表（含账户、姓名、金额的表）
  2、表头要求：
      账户命名，需包含“账”或“帐”
      姓名命名，需包含“姓名”或“户名”
      金额命名，需包含“金额”或“工资”，若存在A金额、B金额等多种，请留下最终要跑批的一列
  3、内容要求：
      账户：须为17/19位账号
      姓名：因柜面系统内个别客户姓名存在空格，所以空格会被保留
      金额：若有多列“XX金额”，请保留一列（精确到小数点后2位）
  
TXT规范：按照柜面批量业务格式
  账号/卡号||姓名|金额|

使用方法：
  1、使用的电脑需安装JDK（提示：OA备份的电脑一般都装有JDK）
  2、直接打开本工具即可使用，绿色免安装

免责声明：
  1、此工具仅提供格式转换，装换后数据准确性请自行校对，本人一概不负责
  2、此工具是本人独立开发，并非农信官方指定工具，如其他联社TXT规范不同，请自行下载源码修改
