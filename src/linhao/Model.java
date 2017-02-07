package linhao;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.HashMap;
import java.util.Map;

import javax.swing.JOptionPane;

import org.apache.log4j.Logger;
import org.apache.log4j.spi.LoggerFactory;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Model {
	
	//private Logger logger = LoggerFactory.getLogger(Model.class);  

    public HSSFWorkbook wb;  
    public HSSFSheet sheet;  
    public HSSFRow row;
    public HSSFCell cell;
    public Integer zhanghao,xingming,jine,kahao,beginRow,endRow,sheetNum;
    //==================================================================================================读取Excel
	public Map<Integer, Map<Integer, String>> readExcel(String filename) throws Exception {//Map<账号，Map<金额，姓名>>
		Map<Integer, Map<Integer, String>> content = new HashMap<Integer, Map<Integer, String>>();
		String ext = filename.substring(filename.lastIndexOf("."));
		String dir = filename.substring(filename.lastIndexOf("\\"));
		try {
			InputStream is = new FileInputStream(filename);
			//================================================================================获取xls
			if (".xls".equals(ext)) {//判断excel文件类型并实例化wb对象
				wb = new HSSFWorkbook(is);
			} else {
				wb = null;
			}
			//JOptionPane.showMessageDialog(null, "替换后"+filename.replace(".xls", ".txt"));
			//================================================================================创建txt
			File f = new File(filename.replace(".xls", ".txt"));  	
            if (f.exists()) {  //根据excel名创建txt文件
            	//JOptionPane.showMessageDialog(null, "txt文件已存在");
            	f.delete();
            	f.createNewFile();
            } else {  
                //System.out.print("文件不存在");  
                f.createNewFile();// 不存在则创建 
            } 
            //================================================================================搜工作表+标志3列+开始行
            updatesheet:for(sheetNum=0;sheetNum<=3;sheetNum++){
            	try {
            		//JOptionPane.showMessageDialog(null, "这是第"+sheetNum+"个工作表");
            		sheet = wb.getSheetAt(sheetNum);// 获取工作表
            		
            		jieshuchazhao:for (int i = 0; i <= 5; i++) {//在前6行搜索账号、姓名、金额所在列，及内容开始行	
                    	row = sheet.getRow(i);
        				if(row==null){
        					//JOptionPane.showMessageDialog(null, "第"+i+"行为空白行");
        					continue;	//跳过空白行
        				}
        				
                    	int jmax = row.getPhysicalNumberOfCells();//第一个非空行的总列数
                    	//JOptionPane.showMessageDialog(null, "第"+i+"行有"+jmax+"列");
         
        				for (int j = 0; j< jmax; j++){//在第一个非空行内逐列搜索
        					cell = row.getCell(j);
        					//JOptionPane.showMessageDialog(null, "j="+j);
        					if(cell==null){
        						//JOptionPane.showMessageDialog(null, "第"+i+"行有"+jmax+"列，第"+j+"列为空白列");
        						continue;//跳过空白列
        					}else if(getValue(cell).contains("账")||getValue(cell).contains("帐")){
        						zhanghao=j;//标志账号
        						//JOptionPane.showMessageDialog(null, "搜索到账户在"+zhanghao+"列，姓名在"+xingming+"列，金额在"+jine+"列，开始行"+beginRow);
        					}else if(getValue(cell).contains("姓名")||getValue(cell).contains("户名")){
        						xingming=j; beginRow=i+1;//标志姓名、开始行
        						//JOptionPane.showMessageDialog(null, "搜索到账户在"+zhanghao+"列，姓名在"+xingming+"列，金额在"+jine+"列，开始行"+beginRow);
        					}else if(getValue(cell).contains("金额")||getValue(cell).contains("工资")){
        						jine=j;//标志金额
        						//JOptionPane.showMessageDialog(null, "搜索到账户在"+zhanghao+"列，姓名在"+xingming+"列，金额在"+jine+"列，开始行"+beginRow);
        					}
        					if(j==jmax){//扫描到最后一列，才能确认账号、金额、姓名在同一行
        						if(zhanghao!=null && xingming !=null && jine != null && beginRow!=null){//如果全找到了，结束搜索
        							//JOptionPane.showMessageDialog(null, "账号、姓名、金额列，开始行全部非空，结束查找");
            						break jieshuchazhao;
            					}
        					}else{
        						continue;
        					}
        				}
        			}

                    if(zhanghao==null || xingming ==null || jine == null || beginRow==null){//无法识别3大列及开始行，提示并推出
                   	 if(sheetNum<=2){//如果是前3个工作表
                   		 //JOptionPane.showMessageDialog(null, "Excel文件："+filename+"第"+sheetNum+"个工作表内都找不到账号、姓名、金额，正在查找下一个工作表");
                   		 continue updatesheet;
                   	 }else{
                   		JOptionPane.showMessageDialog(null, "Excel文件："+filename+"的"+sheetNum+"个工作表内都找不到账号、姓名、金额");
                        System.exit(0);
                   	 }
                    }else{//找出卡列
                    	row = sheet.getRow(beginRow);
                    	int jmax = row.getPhysicalNumberOfCells();
                    	for (int j = 0; j< jmax; j++){//在第一个非空行内逐列搜索
        					cell = row.getCell(j);
        					if(cell==null){//跳过空白列
        						continue;
        					}else	if(getValue(cell).replace(" ", "").replaceAll("\r|\n", "").length()== 19){//找出卡号所在列
        						kahao=j;//标志账号
        						//JOptionPane.showMessageDialog(null, "卡号所在列"+kahao);
        					}
        				}
                    	break updatesheet;
                    }
					//JOptionPane.showMessageDialog(null,dir+"\\"+allFiles[i]);
				} catch (Exception e1) {
					//JOptionPane.showMessageDialog(null, "Excel文件："+filename+"工作表内无法搜索到客户地账号、姓名、金额信息");
					continue updatesheet;
				}
            }
            //JOptionPane.showMessageDialog(null, "确定账户在"+zhanghao+"列，姓名在"+xingming+"列，金额在"+jine+"列，开始行"+beginRow);
            //System.exit(0);
			//================================================================================标志错误行，结束行
			int rowNum = sheet.getLastRowNum(); // 总行数
			
			jieshuhang:for (int i = beginRow; i <= (rowNum+1); i++) {// 从开始行扫描，结束行
				
				try {
					//JOptionPane.showMessageDialog(null, "始行"+beginRow+"，这是"+i+"，总行"+rowNum);
					row = sheet.getRow(i);
					
					if(row==null){//若无单元格，直取结束行
						endRow = i - 1;
						//JOptionPane.showMessageDialog(null, "第"+i+"行为空");
						break jieshuhang;
					}
		            
		            //JOptionPane.showMessageDialog(null, "第"+i+"行，账号类型："+row.getCell(zhanghao).getCellType());
		            
					switch (row.getCell(zhanghao).getCellType()) {//有单元格，则判断账号
					case 0:// 数字
						//JOptionPane.showMessageDialog(null, "第"+i+"行的账号是数字类型");
						if (String.valueOf(row.getCell(zhanghao).getNumericCellValue()).replace(" ", "").replaceAll("\r|\n", "").length() < 17) {// 如果是数字，且不是账号
							endRow = i - 1;
							JOptionPane.showMessageDialog(null,"Excel文件："+filename+"第"+sheetNum+"个工作表，第"+(i+1)+"行数字账号位数不对："+String.valueOf(row.getCell(zhanghao).getNumericCellValue()));
							System.exit(0);
							break jieshuhang;
						}else{
							break;
						}
					case 1:// 中文
						//JOptionPane.showMessageDialog(null, "文本"+i);
						if (String.valueOf(row.getCell(zhanghao).getStringCellValue()).replace(" ", "").replaceAll("\r|\n", "").length() != 17 & String.valueOf(row.getCell(zhanghao).getStringCellValue()).replace(" ", "").replaceAll("\r|\n", "").length() != 19) {
							endRow = i - 1;
							JOptionPane.showMessageDialog(null, "Excel文件："+filename+"第"+sheetNum+"个工作表，第"+(i+1)+"行文本账号位数不对：" + String.valueOf(row.getCell(zhanghao).getStringCellValue()));
							System.exit(0);
							break jieshuhang;
						}else{
							break;
						}
					case 3:// 3空值HSSFCell.CELL_TYPE_BLANK
						try {
							if(getValue(row.getCell(kahao)).replace(" ", "").replaceAll("\r|\n", "").length()== 19){//如果有卡号，继续扫描
								break;
							}else{
								//JOptionPane.showMessageDialog(null, "Excel文件："+filename+"第"+sheetNum+"个工作表，第"+(i+1) + "行账号列为空，卡号列也为空，结束扫描");
								endRow = i - 1;
								break jieshuhang;
							}
						} catch (Exception e1) {//加入异常捕捉
							//JOptionPane.showMessageDialog(null,"Excel文件："+filename+"第"+sheetNum+"个工作表，第"+(i+1)+ "行账号列为空，不存在卡号列，结束扫描");
							endRow = i - 1;
							break jieshuhang;
						}
					}
				} catch (Exception e1) {
					//JOptionPane.showMessageDialog(null, "Excel文件："+filename+"第"+sheetNum+"个工作表，第"+(i+1)+ "行无法获取账号类型");	
					endRow = i - 1;
					break jieshuhang;
				}
			}
			//JOptionPane.showMessageDialog(null, "错误扫描结束，最终结束行"+endRow+"正在判断该行是否正常结束行");
			//================================================================================判断结束行是否最后一行
			if(endRow<rowNum){
				row = sheet.getRow(endRow+1);
				try {
					if(row==null){
						//JOptionPane.showMessageDialog(null, "正常结束：整行为空");
					}else if(row.getCell(zhanghao).getCellType()==3){
						//JOptionPane.showMessageDialog(null, "正常结束：账号列卡号列均为空");
					}else{
						JOptionPane.showMessageDialog(null, "Excel文件："+filename+"第"+sheetNum+"个工作表，第"+(endRow+1)+"行账号内容异常，请检查");
						System.exit(0);
					}
				} catch (Exception e1) {
					//JOptionPane.showMessageDialog(null, "Excel文件："+filename+"第"+sheetNum+"个工作表，第"+(endRow+1)+"行无法获取账号列类型");
				}
			}
			//JOptionPane.showMessageDialog(null, "总行数："+rowNum+"，开始行："+beginRow+"，停止行："+endRow);
			//================================================================================读excel写txt
			BufferedWriter output = new BufferedWriter(new FileWriter(f));
			Double total = 0.00;//金额、总金额
			DecimalFormat liangwei = new DecimalFormat( "0.00 ");
			
			for (int i = beginRow; i <= endRow; i++) {//从开始行开始
				row = sheet.getRow(i); //获取行
				if(i==endRow){//如果是最后一行
					if(getValue(row.getCell(zhanghao))=="单元格为空"){//如果没账号就用卡号
						total = total +Double.valueOf(getValue(row.getCell(jine)));
						//JOptionPane.showMessageDialog(null, "total:"+total);
						
						output.write(getValue(row.getCell(kahao))+"||"+getValue(row.getCell(xingming))+"|"+getValue(row.getCell(jine))+"|");
					}else{//有账号默认用账号
						total = total +Double.valueOf(getValue(row.getCell(jine)));
						//JOptionPane.showMessageDialog(null, "total:"+total);
						
						output.write(getValue(row.getCell(zhanghao))+"||"+getValue(row.getCell(xingming))+"|"+getValue(row.getCell(jine))+"|");
					}
				}else{//除了最后一行，后面的全部添加换行符
					if(getValue(row.getCell(zhanghao))=="单元格为空"){
						total = total +Double.valueOf(getValue(row.getCell(jine)));
						output.write(getValue(row.getCell(kahao))+"||"+getValue(row.getCell(xingming))+"|"+getValue(row.getCell(jine))+"|\r\n");	
					}else{
						total = total +Double.valueOf(getValue(row.getCell(jine)));
						output.write(getValue(row.getCell(zhanghao))+"||"+getValue(row.getCell(xingming))+"|"+getValue(row.getCell(jine))+"|\r\n");	
					}
				}
				//System.out.println(getValue(row.getCell(zhanghao))+"||"+getValue(row.getCell(xingming))+"|"+getValue(row.getCell(jine))+"|\r\n");
			}
			output.close();
			//JOptionPane.showMessageDialog(null, "最后账户"+getValue(sheet.getRow(rowNum).getCell(twoCol)));
			is.close();

			JOptionPane.showMessageDialog(null, "Excel文件："+filename+"的第"+sheetNum+"个工作表，有"+(endRow-beginRow+1)+"笔，合计"+liangwei.format(total)+"元");
		} catch (FileNotFoundException e) {
			// 找不到文件
		} catch (IOException e) {
			// 读取异常
		}
		return content;
	}
	
	public String getValue(HSSFCell cell) {//获取单元格的值
		java.text.DecimalFormat zhanghao = new java.text.DecimalFormat("########");
		DecimalFormat   jine  = new DecimalFormat("0.00");//十进制格式
		
		switch (cell.getCellType()) {
		case 0: // 0数字HSSFCell.CELL_TYPE_NUMERIC：账号，金额，日期
			//JOptionPane.showMessageDialog(null, cell.getNumericCellValue());break;
			if(String.valueOf(cell.getNumericCellValue()).replace(" ", "").length()>=17){//如果是账号
				return  String.valueOf(zhanghao.format(cell.getNumericCellValue())).replace(" ", "").replaceAll("\r|\n", "");
				//return  String.valueOf(cell.getNumericCellValue());
			}else{//如果是金额
				return  String.valueOf(jine.format(cell.getNumericCellValue())).replace(" ", "").replaceAll("\r|\n", "");
			}
		case 1:// 1中文HSSFCell.CELL_TYPE_STRING：姓名
			//JOptionPane.showMessageDialog(null, cell.getStringCellValue());break;
			return String.valueOf(cell.getStringCellValue()).replaceAll("\r|\n", "");
		case 2:// 2公式HSSFCell.CELL_TYPE_FORMULA
				// JOptionPane.showMessageDialog(null, "单元格插入公式");break;
			try {
				return String.valueOf(jine.format(cell.getNumericCellValue())).replace(" ", "").replaceAll("\r|\n", "");
			} catch (IllegalStateException e) {
				return String.valueOf(jine.format(cell.getRichStringCellValue())).replace(" ", "").replaceAll("\r|\n", "");
			}
		case 3:// 3空值HSSFCell.CELL_TYPE_BLANK
			//JOptionPane.showMessageDialog(null, "单元格为空");break;
			return "单元格为空";
		case 4:// 4真假HSSFCell.CELL_TYPE_BOOLEAN：true/false
			//JOptionPane.showMessageDialog(null, cell.getBooleanCellValue()); break;
			return String.valueOf(cell.getBooleanCellValue()).replace(" ", "").replaceAll("\r|\n", "");
		case 5:// 5错误HSSFCell.CELL_TYPE_ERROR：#N/A
			//JOptionPane.showMessageDialog(null, "单元格内容错误"); break;
			return "单元格内容错误";
		}
		return "没有匹配值类型";
	}

}
