package cc.watchers.ewio.xlsxUtils;


import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author http://www.2cto.com/kf/201605/510933.html 
 */

public class XlsxWrite {
	
	
	/**
	 * 写入数值 
	 */
	public static void writeExcel(OutputStream os, String excelExtName, List<List<List<String>>> data){
		
		if(data==null) {
			return;
		}			
		Workbook wb = null;
		
        if ("xls".equals(excelExtName)) {
            wb = new HSSFWorkbook();
        } else if ("xlsx".equals(excelExtName)) {
            wb = new XSSFWorkbook();
        } else {
        	return;
        	//throw new Exception("当前文件不是excel文件");
        }
        
		try {
			int i,j,k;
			i=j=k=0;
	        for (List<List<String>> sdata : data) {
	        	Sheet sheet = wb.createSheet(String.valueOf(i));
	        	j=0;
	        	for(List<String> ldata:sdata) {
	        		Row row = sheet.createRow(j);
	        		k=0;
	        		for(String rdata:ldata) {
	        			Cell cell = row.createCell(k);
	        			cell.setCellValue(rdata);
	        			k++;
	        		}
	        		j++;
	        	}
	        	i++;
	        }
	        wb.write(os);
			
		} catch (Exception e) {
	        e.printStackTrace();
	    } finally {
	        if (wb != null) {
	            try {
					wb.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
	        }
	    }
		
	}
	
	
	public static void writeExcel(OutputStream os, String excelExtName, Map<String, List<List<String>>> data) throws IOException{
	    Workbook wb = null;
	    try {
	        if ("xls".equals(excelExtName)) {
	            wb = new HSSFWorkbook();
	        } else if ("xlsx".equals(excelExtName)) {
	            wb = new XSSFWorkbook();
	        } else {
	            throw new Exception("当前文件不是excel文件");
	        }
	        for (String sheetName : data.keySet()) {
	            Sheet sheet = wb.createSheet(sheetName);
	            List<List<String>> rowList = data.get(sheetName);
	            for (int i = 0; i < rowList.size(); i++) {
	                List<String> cellList = rowList.get(i);
	                Row row = sheet.createRow(i);
	                for (int j = 0; j < cellList.size(); j++) {
	                    Cell cell = row.createCell(j);
	                    cell.setCellValue(cellList.get(j));
	                }
	            }
	        }
	        wb.write(os);
	    } catch (Exception e) {
	        e.printStackTrace();
	    } finally {
	        if (wb != null) {
	            wb.close();
	        }
	    }
	}

}
