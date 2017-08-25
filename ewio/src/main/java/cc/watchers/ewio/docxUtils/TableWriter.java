package cc.nitsc.ewio.docxUtils;

import java.util.List;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;



public class TableWriter {
	
	// 从表后追加数据，第startRowNumber行开始重写
	static public void writeRows(XWPFTable table, List<List<String>> data, int startRowNumber) {

		if (data == null || data.size() == 0) {
			return;
		}
		int addSize = data.size();
		int expectSize = startRowNumber + addSize - 1;
		int nowSize = table.getNumberOfRows();
		if (nowSize - 1 < expectSize) {
			addTableRowsSpace(table, expectSize - nowSize+1);
		}
		// 更新表格数据
		List<XWPFTableRow> rows;
		List<XWPFTableCell> rowCells;
		rows = table.getRows();
		
		int rIndex=startRowNumber;
		
		//写xx行		
		for(int i=0;i<data.size();i++,rIndex++) {
			XWPFTableRow row = rows.get(rIndex);
			
			List<String> rowData = data.get(i);
			rowCells = row.getTableCells();//一行cell
			
			if(rowData.size()!=rowCells.size()) {
				System.out.println("警告:数据与表格格式不匹配");
				System.out.println(rowData);
				System.out.println("rData:"+rowData.size());
				System.out.println("cells:"+rowCells.size());
			}
			
			//写xx格
			for(int j=0;j<rowData.size();j++) {
				XWPFTableCell cell = rowCells.get(j);
				CTP ctp = CTP.Factory.newInstance();
				XWPFParagraph p = new XWPFParagraph(ctp, cell);
				//p.setAlignment(ParagraphAlignment.CENTER);
				p.setAlignment(ParagraphAlignment.LEFT);
				XWPFRun run = p.createRun();
				run.setFontSize(11);
				
				//run.setBold(true);//加粗
				if( rIndex>1 && rIndex<5 && data.size()==35 && rowData.size()==2)
				{
					System.out.println("TableWriter特定格式加粗体");
					run.setBold(true);
				}
				run.setText(rowData.get(j));
				//...
				CTRPr rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
				CTFonts fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
				fonts.setAscii("等线");
				fonts.setEastAsia("等线");
				fonts.setHAnsi("等线");
				cell.setParagraph(p);
			}
			
		}
	}
	
	// 表后追加n行
	public static void addTableRowsSpace(XWPFTable table, int addSize) {
		for (int i = 0; i < addSize; i++) {
			table.createRow();// 加一行
		}
	}

	
	/**
	 * @Description: 跨行合并
	 * @see http://stackoverflow.com/questions/24907541/row-span-with-xwpftable
	 */
	static public void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
		for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
			XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
			if (rowIndex == fromRow) {
				// The first merged cell is set with RESTART merge value
				cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
			} else {
				// Cells which join (merge) the first one, are set with CONTINUE
				cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
			}
		}
	}

	
	
	// word跨列合并单元格
	@SuppressWarnings("unused")
	static private void mergeCellsHorizontal(XWPFTable table, int row, int fromCell, int toCell) {
		for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
			XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
			if (cellIndex == fromCell) {
				// The first merged cell is set with RESTART merge value
				cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
			} else {
				// Cells which join (merge) the first one, are set with CONTINUE
				cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
			}
		}
	}

	//设置字体
	@SuppressWarnings("unused")
	static public void getParagraph(XWPFTableCell cell, String cellText) {
		CTP ctp = CTP.Factory.newInstance();
		XWPFParagraph p = new XWPFParagraph(ctp, cell);
		p.setAlignment(ParagraphAlignment.CENTER);
		XWPFRun run = p.createRun();
		run.setText(cellText);
		CTRPr rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
		CTFonts fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
		fonts.setAscii("仿宋");
		fonts.setEastAsia("仿宋");
		fonts.setHAnsi("仿宋");
		cell.setParagraph(p);
	}

}


/*				


//每个单元格子，可能有"多个"文字段，先移除
paras = cell.getParagraphs();
for (XWPFParagraph para : paras) {
	List<XWPFRun> runs = para.getRuns();
	for (int ttt = runs.size(); ttt > 0; ttt--) {
		para.removeRun(0);
	}
	XWPFRun newInfRun = para.createRun();
	// dfs.setDocxFontStyle(newInfRun);
	newInfRun.setText(rowData.get(j));
	
}

*/
