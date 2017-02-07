# ExcelToTxt
  概述：将一个/多个客户的Excel转换成柜面批量业务的TXT

Excel规范：已按照客户送来的格式规范进行最大限度兼容
  1、仅支持xls格式
  2、存储客户信息（账户、姓名、金额）的工作表必须放在第一位
  3、表头（包含账户、姓名、金额）的一行在前6行
  4、账户规则：命名需包含“账”或“帐”，内容为17/19位账号（自动去除空格）
  5、姓名规则：命名需包含“姓名”或“户名”，默认保留空格（柜面系统内有客户姓名存在空格）
  6、金额规则：命名需包含“金额”或“工资”，若有多列“XX金额”，请保留一列（精确到小数点后2位）
TXT规范：按照柜面批量业务格式
  账号/卡号||姓名|金额|
  
功能：
  1、不用删除客户Excel内无关列
  2、支持文件夹内所有Excel批量转换
  3、账号位数错误会具体提示“某Excel的xx工作表的xx行错误”
  4、
