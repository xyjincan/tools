package cc.watchers.ewio.xlsxUtils;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author http://www.2cto.com/kf/201605/510933.html 
 */

public class XlsxRead {
	
	/**
	 * 读取数值 
	 */
	public static List<List<List<String>>> readExcelWithoutTitle(String filepath){
	    
		String fileType = filepath.substring(filepath.lastIndexOf(".") + 1, filepath.length());
	    InputStream is = null;
	    Workbook wb = null;
	    try {
	        is = new FileInputStream(filepath);
	        if (fileType.equals("xls")) {
	            wb = new HSSFWorkbook(is);
	        } else if (fileType.equals("xlsx")) {
	            wb = new XSSFWorkbook(is);
	        } else {
	            //throw new Exception("读取的不是excel文件");
	        }
	         
	        List<List<List<String>>> result = new ArrayList<List<List<String>>>();//对应excel文件
	         
	        int sheetSize = wb.getNumberOfSheets();
	        for (int i = 0; i < sheetSize; i++) {//遍历sheet页
	            Sheet sheet = wb.getSheetAt(i);
	            List<List<String>> sheetList = new ArrayList<List<String>>();//对应sheet页
	             
	            int rowSize = sheet.getLastRowNum() + 1;
	            for (int j = 0; j < rowSize; j++) {//遍历行
	                Row row = sheet.getRow(j);
	                if (row == null) {//略过空行
	                    continue;
	                }
	                int cellSize = row.getLastCellNum();//行中有多少个单元格，也就是有多少列
	                List<String> rowList = new ArrayList<String>();//对应一个数据行
	                for (int k = 0; k < cellSize; k++) {
	                    Cell cell = row.getCell(k);
	                    String value = null;
	                    if (cell != null) {
	                        value = cell.toString();
	                    }
	                    rowList.add(value);
	                }
	                sheetList.add(rowList);
	            }
	            result.add(sheetList);
	        }
	        
	        return result;
	        
	    } catch (FileNotFoundException e) {
	    } catch (IOException e) {
		} finally {
	        if (wb != null) {
	            try {
					wb.close();
				} catch (IOException e) {
				}
	        }
	        if (is != null) {
	            try {
					is.close();
				} catch (IOException e) {
				}
	        }
	    }
		return null;
	}

}
