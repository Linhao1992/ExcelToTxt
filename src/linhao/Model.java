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
    //==================================================================================================��ȡExcel
	public Map<Integer, Map<Integer, String>> readExcel(String filename) throws Exception {//Map<�˺ţ�Map<������>>
		Map<Integer, Map<Integer, String>> content = new HashMap<Integer, Map<Integer, String>>();
		String ext = filename.substring(filename.lastIndexOf("."));
		String dir = filename.substring(filename.lastIndexOf("\\"));
		try {
			InputStream is = new FileInputStream(filename);
			//================================================================================��ȡxls
			if (".xls".equals(ext)) {//�ж�excel�ļ����Ͳ�ʵ����wb����
				wb = new HSSFWorkbook(is);
			} else {
				wb = null;
			}
			//JOptionPane.showMessageDialog(null, "�滻��"+filename.replace(".xls", ".txt"));
			//================================================================================����txt
			File f = new File(filename.replace(".xls", ".txt"));  	
            if (f.exists()) {  //����excel������txt�ļ�
            	//JOptionPane.showMessageDialog(null, "txt�ļ��Ѵ���");
            	f.delete();
            	f.createNewFile();
            } else {  
                //System.out.print("�ļ�������");  
                f.createNewFile();// �������򴴽� 
            } 
            //================================================================================�ѹ�����+��־3��+��ʼ��
            updatesheet:for(sheetNum=0;sheetNum<=3;sheetNum++){
            	try {
            		//JOptionPane.showMessageDialog(null, "���ǵ�"+sheetNum+"��������");
            		sheet = wb.getSheetAt(sheetNum);// ��ȡ������
            		
            		jieshuchazhao:for (int i = 0; i <= 5; i++) {//��ǰ6�������˺š���������������У������ݿ�ʼ��	
                    	row = sheet.getRow(i);
        				if(row==null){
        					//JOptionPane.showMessageDialog(null, "��"+i+"��Ϊ�հ���");
        					continue;	//�����հ���
        				}
        				
                    	int jmax = row.getPhysicalNumberOfCells();//��һ���ǿ��е�������
                    	//JOptionPane.showMessageDialog(null, "��"+i+"����"+jmax+"��");
         
        				for (int j = 0; j< jmax; j++){//�ڵ�һ���ǿ�������������
        					cell = row.getCell(j);
        					//JOptionPane.showMessageDialog(null, "j="+j);
        					if(cell==null){
        						//JOptionPane.showMessageDialog(null, "��"+i+"����"+jmax+"�У���"+j+"��Ϊ�հ���");
        						continue;//�����հ���
        					}else if(getValue(cell).contains("��")||getValue(cell).contains("��")){
        						zhanghao=j;//��־�˺�
        						//JOptionPane.showMessageDialog(null, "�������˻���"+zhanghao+"�У�������"+xingming+"�У������"+jine+"�У���ʼ��"+beginRow);
        					}else if(getValue(cell).contains("����")||getValue(cell).contains("����")){
        						xingming=j; beginRow=i+1;//��־��������ʼ��
        						//JOptionPane.showMessageDialog(null, "�������˻���"+zhanghao+"�У�������"+xingming+"�У������"+jine+"�У���ʼ��"+beginRow);
        					}else if(getValue(cell).contains("���")||getValue(cell).contains("����")){
        						jine=j;//��־���
        						//JOptionPane.showMessageDialog(null, "�������˻���"+zhanghao+"�У�������"+xingming+"�У������"+jine+"�У���ʼ��"+beginRow);
        					}
        					if(j==jmax){//ɨ�赽���һ�У�����ȷ���˺š���������ͬһ��
        						if(zhanghao!=null && xingming !=null && jine != null && beginRow!=null){//���ȫ�ҵ��ˣ���������
        							//JOptionPane.showMessageDialog(null, "�˺š�����������У���ʼ��ȫ���ǿգ���������");
            						break jieshuchazhao;
            					}
        					}else{
        						continue;
        					}
        				}
        			}

                    if(zhanghao==null || xingming ==null || jine == null || beginRow==null){//�޷�ʶ��3���м���ʼ�У���ʾ���Ƴ�
                   	 if(sheetNum<=2){//�����ǰ3��������
                   		 //JOptionPane.showMessageDialog(null, "Excel�ļ���"+filename+"��"+sheetNum+"���������ڶ��Ҳ����˺š������������ڲ�����һ��������");
                   		 continue updatesheet;
                   	 }else{
                   		JOptionPane.showMessageDialog(null, "Excel�ļ���"+filename+"��"+sheetNum+"���������ڶ��Ҳ����˺š����������");
                        System.exit(0);
                   	 }
                    }else{//�ҳ�����
                    	row = sheet.getRow(beginRow);
                    	int jmax = row.getPhysicalNumberOfCells();
                    	for (int j = 0; j< jmax; j++){//�ڵ�һ���ǿ�������������
        					cell = row.getCell(j);
        					if(cell==null){//�����հ���
        						continue;
        					}else	if(getValue(cell).replace(" ", "").replaceAll("\r|\n", "").length()== 19){//�ҳ�����������
        						kahao=j;//��־�˺�
        						//JOptionPane.showMessageDialog(null, "����������"+kahao);
        					}
        				}
                    	break updatesheet;
                    }
					//JOptionPane.showMessageDialog(null,dir+"\\"+allFiles[i]);
				} catch (Exception e1) {
					//JOptionPane.showMessageDialog(null, "Excel�ļ���"+filename+"���������޷��������ͻ����˺š������������Ϣ");
					continue updatesheet;
				}
            }
            //JOptionPane.showMessageDialog(null, "ȷ���˻���"+zhanghao+"�У�������"+xingming+"�У������"+jine+"�У���ʼ��"+beginRow);
            //System.exit(0);
			//================================================================================��־�����У�������
			int rowNum = sheet.getLastRowNum(); // ������
			
			jieshuhang:for (int i = beginRow; i <= (rowNum+1); i++) {// �ӿ�ʼ��ɨ�裬������
				
				try {
					//JOptionPane.showMessageDialog(null, "ʼ��"+beginRow+"������"+i+"������"+rowNum);
					row = sheet.getRow(i);
					
					if(row==null){//���޵�Ԫ��ֱȡ������
						endRow = i - 1;
						//JOptionPane.showMessageDialog(null, "��"+i+"��Ϊ��");
						break jieshuhang;
					}
		            
		            //JOptionPane.showMessageDialog(null, "��"+i+"�У��˺����ͣ�"+row.getCell(zhanghao).getCellType());
		            
					switch (row.getCell(zhanghao).getCellType()) {//�е�Ԫ�����ж��˺�
					case 0:// ����
						//JOptionPane.showMessageDialog(null, "��"+i+"�е��˺�����������");
						if (String.valueOf(row.getCell(zhanghao).getNumericCellValue()).replace(" ", "").replaceAll("\r|\n", "").length() < 17) {// ��������֣��Ҳ����˺�
							endRow = i - 1;
							JOptionPane.showMessageDialog(null,"Excel�ļ���"+filename+"��"+sheetNum+"����������"+(i+1)+"�������˺�λ�����ԣ�"+String.valueOf(row.getCell(zhanghao).getNumericCellValue()));
							System.exit(0);
							break jieshuhang;
						}else{
							break;
						}
					case 1:// ����
						//JOptionPane.showMessageDialog(null, "�ı�"+i);
						if (String.valueOf(row.getCell(zhanghao).getStringCellValue()).replace(" ", "").replaceAll("\r|\n", "").length() != 17 & String.valueOf(row.getCell(zhanghao).getStringCellValue()).replace(" ", "").replaceAll("\r|\n", "").length() != 19) {
							endRow = i - 1;
							JOptionPane.showMessageDialog(null, "Excel�ļ���"+filename+"��"+sheetNum+"����������"+(i+1)+"���ı��˺�λ�����ԣ�" + String.valueOf(row.getCell(zhanghao).getStringCellValue()));
							System.exit(0);
							break jieshuhang;
						}else{
							break;
						}
					case 3:// 3��ֵHSSFCell.CELL_TYPE_BLANK
						try {
							if(getValue(row.getCell(kahao)).replace(" ", "").replaceAll("\r|\n", "").length()== 19){//����п��ţ�����ɨ��
								break;
							}else{
								//JOptionPane.showMessageDialog(null, "Excel�ļ���"+filename+"��"+sheetNum+"����������"+(i+1) + "���˺���Ϊ�գ�������ҲΪ�գ�����ɨ��");
								endRow = i - 1;
								break jieshuhang;
							}
						} catch (Exception e1) {//�����쳣��׽
							//JOptionPane.showMessageDialog(null,"Excel�ļ���"+filename+"��"+sheetNum+"����������"+(i+1)+ "���˺���Ϊ�գ������ڿ����У�����ɨ��");
							endRow = i - 1;
							break jieshuhang;
						}
					}
				} catch (Exception e1) {
					//JOptionPane.showMessageDialog(null, "Excel�ļ���"+filename+"��"+sheetNum+"����������"+(i+1)+ "���޷���ȡ�˺�����");	
					endRow = i - 1;
					break jieshuhang;
				}
			}
			//JOptionPane.showMessageDialog(null, "����ɨ����������ս�����"+endRow+"�����жϸ����Ƿ�����������");
			//================================================================================�жϽ������Ƿ����һ��
			if(endRow<rowNum){
				row = sheet.getRow(endRow+1);
				try {
					if(row==null){
						//JOptionPane.showMessageDialog(null, "��������������Ϊ��");
					}else if(row.getCell(zhanghao).getCellType()==3){
						//JOptionPane.showMessageDialog(null, "�����������˺��п����о�Ϊ��");
					}else{
						JOptionPane.showMessageDialog(null, "Excel�ļ���"+filename+"��"+sheetNum+"����������"+(endRow+1)+"���˺������쳣������");
						System.exit(0);
					}
				} catch (Exception e1) {
					//JOptionPane.showMessageDialog(null, "Excel�ļ���"+filename+"��"+sheetNum+"����������"+(endRow+1)+"���޷���ȡ�˺�������");
				}
			}
			//JOptionPane.showMessageDialog(null, "��������"+rowNum+"����ʼ�У�"+beginRow+"��ֹͣ�У�"+endRow);
			//================================================================================��excelдtxt
			BufferedWriter output = new BufferedWriter(new FileWriter(f));
			Double total = 0.00;//���ܽ��
			DecimalFormat liangwei = new DecimalFormat( "0.00 ");
			
			for (int i = beginRow; i <= endRow; i++) {//�ӿ�ʼ�п�ʼ
				row = sheet.getRow(i); //��ȡ��
				if(i==endRow){//��������һ��
					if(getValue(row.getCell(zhanghao))=="��Ԫ��Ϊ��"){//���û�˺ž��ÿ���
						total = total +Double.valueOf(getValue(row.getCell(jine)));
						//JOptionPane.showMessageDialog(null, "total:"+total);
						
						output.write(getValue(row.getCell(kahao))+"||"+getValue(row.getCell(xingming))+"|"+getValue(row.getCell(jine))+"|");
					}else{//���˺�Ĭ�����˺�
						total = total +Double.valueOf(getValue(row.getCell(jine)));
						//JOptionPane.showMessageDialog(null, "total:"+total);
						
						output.write(getValue(row.getCell(zhanghao))+"||"+getValue(row.getCell(xingming))+"|"+getValue(row.getCell(jine))+"|");
					}
				}else{//�������һ�У������ȫ����ӻ��з�
					if(getValue(row.getCell(zhanghao))=="��Ԫ��Ϊ��"){
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
			//JOptionPane.showMessageDialog(null, "����˻�"+getValue(sheet.getRow(rowNum).getCell(twoCol)));
			is.close();

			JOptionPane.showMessageDialog(null, "Excel�ļ���"+filename+"�ĵ�"+sheetNum+"����������"+(endRow-beginRow+1)+"�ʣ��ϼ�"+liangwei.format(total)+"Ԫ");
		} catch (FileNotFoundException e) {
			// �Ҳ����ļ�
		} catch (IOException e) {
			// ��ȡ�쳣
		}
		return content;
	}
	
	public String getValue(HSSFCell cell) {//��ȡ��Ԫ���ֵ
		java.text.DecimalFormat zhanghao = new java.text.DecimalFormat("########");
		DecimalFormat   jine  = new DecimalFormat("0.00");//ʮ���Ƹ�ʽ
		
		switch (cell.getCellType()) {
		case 0: // 0����HSSFCell.CELL_TYPE_NUMERIC���˺ţ�������
			//JOptionPane.showMessageDialog(null, cell.getNumericCellValue());break;
			if(String.valueOf(cell.getNumericCellValue()).replace(" ", "").length()>=17){//������˺�
				return  String.valueOf(zhanghao.format(cell.getNumericCellValue())).replace(" ", "").replaceAll("\r|\n", "");
				//return  String.valueOf(cell.getNumericCellValue());
			}else{//����ǽ��
				return  String.valueOf(jine.format(cell.getNumericCellValue())).replace(" ", "").replaceAll("\r|\n", "");
			}
		case 1:// 1����HSSFCell.CELL_TYPE_STRING������
			//JOptionPane.showMessageDialog(null, cell.getStringCellValue());break;
			return String.valueOf(cell.getStringCellValue()).replaceAll("\r|\n", "");
		case 2:// 2��ʽHSSFCell.CELL_TYPE_FORMULA
				// JOptionPane.showMessageDialog(null, "��Ԫ����빫ʽ");break;
			try {
				return String.valueOf(jine.format(cell.getNumericCellValue())).replace(" ", "").replaceAll("\r|\n", "");
			} catch (IllegalStateException e) {
				return String.valueOf(jine.format(cell.getRichStringCellValue())).replace(" ", "").replaceAll("\r|\n", "");
			}
		case 3:// 3��ֵHSSFCell.CELL_TYPE_BLANK
			//JOptionPane.showMessageDialog(null, "��Ԫ��Ϊ��");break;
			return "��Ԫ��Ϊ��";
		case 4:// 4���HSSFCell.CELL_TYPE_BOOLEAN��true/false
			//JOptionPane.showMessageDialog(null, cell.getBooleanCellValue()); break;
			return String.valueOf(cell.getBooleanCellValue()).replace(" ", "").replaceAll("\r|\n", "");
		case 5:// 5����HSSFCell.CELL_TYPE_ERROR��#N/A
			//JOptionPane.showMessageDialog(null, "��Ԫ�����ݴ���"); break;
			return "��Ԫ�����ݴ���";
		}
		return "û��ƥ��ֵ����";
	}

}
