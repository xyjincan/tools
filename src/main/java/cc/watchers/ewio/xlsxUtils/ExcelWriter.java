package cc.watchers.ewio.xlsxUtils;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 
 * 打开一个excel对象，当所有的数据写入完成后才能关闭
 * 
 */

public class ExcelWriter {
	
	private Workbook wb = null;
	private String resultFilePath;
	
	/**
	 * 
	 * 使用指定模板填入需要的数值到新文件
	 * 模板与新文件相同时，新数据替换旧数据
	 */
	public ExcelWriter(String template,String resultFilePath){
		String fileType = resultFilePath.substring(resultFilePath.lastIndexOf(".") + 1, resultFilePath.length());
		FileInputStream is = null;
		try {
	        if (fileType.equals("xls")) {
	    		is = new FileInputStream(template);
	            wb = new HSSFWorkbook(is);
	        } else if (fileType.equals("xlsx")) {
	    		is = new FileInputStream(template);
	            wb = new XSSFWorkbook(is);
	        } else {
	            new Exception("读取的不是excel文件").printStackTrace();;
	        }
		}catch(Exception e) {
		}
		this.resultFilePath = resultFilePath;
	}
	
	
	/**
	 * 
	 * 连续保存n个sheet
	 * data 数据源
	 * position 保存数据起始坐标
	 * 
	 */
    
	public void writeNS(List<List<List<String>>> data,ExcelPosition position) {
		
		if(wb==null || data==null || position==null) {
			return;
		}
		int sheetSize = data.size()+position.s;
		if(sheetSize<0) {
			return;
		}
		//
        for (int i=position.s; i < sheetSize; i++) {//遍历sheet页
            List<List<String>> sheetList = data.get(i-position.s);//
            writeNL(sheetList,new ExcelPosition(i, position.l, position.c));
        }
	}
	
	/**
	 * 
	 * 连续保存n行数据
	 * data 数据源
	 * position 保存数据起始坐标
	 * 
	 */
	public void writeNL(List<List<String>> data,ExcelPosition position) {
		
		if(wb==null || data==null || position==null) {
			return;
		}
		int sheetSize = position.s;
		if(sheetSize<0) {
			return;
		}
		
		if(wb.getNumberOfSheets()<=sheetSize) {
			int add = sheetSize-wb.getNumberOfSheets();
			for(int i=0;i<=add;i++) {
				wb.createSheet();//没有则新建，保证数据保存到了需要的地方
			}
		}
		
		int i,j,k;
		i=position.s;
        Sheet sheet = wb.getSheetAt(i);
        int rowSize = data.size() + position.l;
        for (j=position.l; j < rowSize; j++) {//依次保存每行
            Row row = sheet.getRow(j);
            List<String> rowList = data.get(j-position.l);
            if (row == null) {
            	row = sheet.createRow(j);
            }
            int cellSize = rowList.size()+position.c;
            for (k=position.c; k < cellSize; k++) {//依次保存每个单元格
                Cell cell = row.getCell(k);
                if(cell == null) {
                	cell = row.createCell(k);
                }
                String cellData = rowList.get(k-position.c);
                //强行使用异常判别 保存对应的数值类型,,,
                try {
                	cell.setCellValue(Integer.valueOf(cellData));
                }catch (Exception ei) {
                    try {
                    	cell.setCellValue(Double.valueOf(cellData));	
                    }catch (Exception ed) {
                    	cell.setCellValue(cellData);	
					}
				}
            }
        }
	}
	
	/**
	 * 动态以某名字在Excel最后面追加sheet
	 */
	public void createSheets(List<String> names) {
		
		if(names==null || names.size() == 0) {
			return;
		}
		for(int i=0;i < names.size();i++) {
			wb.createSheet(names.get(i));
		}
	}
	
	
	
	/**
	 * 将数据保存到目标文件
	 */
	public void close(){
		try {
			OutputStream os = new FileOutputStream(resultFilePath);
			wb.write(os);
			wb.close();
			os.close();
			wb=null;
		}catch(IOException ioe) {
			new Exception("写入文件失败："+resultFilePath).printStackTrace();
		}
	}
}
