package cc.watchers.ewio.docxUtils;

import java.io.Closeable;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 * 
 * 只支持了docx格式
 * 占位符${xxxxss}
 * xxxx标识文档中需要替换的定位，ss表示该位置需要的字体大小，默认16
 * 
 * 参考：http://blog.sina.com.cn/s/blog_87df085f0102w17v.html
 */

public class TemplateWordWriter {
	
	public void TemplateWriter(String template,String resultFilePath,Map<String, String> params){

		try {
			InputStream is = new FileInputStream(template);
			XWPFDocument doc = new XWPFDocument(is);
			// 替换文档里面的变量
			this.replaceInDocument(doc, params);
			// 替换表格里面的变量
			this.replaceInTable(doc, params);
			OutputStream os = new FileOutputStream(resultFilePath);
			doc.write(os);
			close(os);
			close(is);
		} catch (IOException e) {
			new IOException(template+" 写入 "+resultFilePath+" 失败").printStackTrace();
		}
	}
	
	private void replaceInDocument(XWPFDocument docx, Map<String, String> params) {
		
		List<XWPFParagraph> paragraphList = docx.getParagraphs();//获取word标题列表
		for(XWPFParagraph para:paragraphList) {
			replaceText(para,params);
		}
	}

	private void replaceInTable(XWPFDocument doc, Map<String, String> params) {

		Iterator<XWPFTable> iterator = doc.getTablesIterator();
		XWPFTable table;
		List<XWPFTableRow> rows;
		List<XWPFTableCell> cells;
		List<XWPFParagraph> paras;
		while (iterator.hasNext()) {
			table = iterator.next();
			rows = table.getRows();
			for (XWPFTableRow row : rows) {
				cells = row.getTableCells();
				for (XWPFTableCell cell : cells) {
					paras = cell.getParagraphs();
					for (XWPFParagraph para : paras) {
						replaceText(para,params);
					}
				}
			}
		}
	}

	private void replaceText(XWPFParagraph para,Map<String, String> params) {
		
		String paraText = para.getText();
		if(paraText.trim().length()==0){	
			return;
		}
		String runText = paraText;
		if (this.matcher(paraText).find()) {
			//替换过程中，如果发现字体大小设置信息，会更新到fontSize中
			runText = matcherReplace(paraText, params);
			List<XWPFRun> runs = para.getRuns();
			DocxFontStyle dfs = new DocxFontStyle(runs.get(0),fontSize);
			for(int i=runs.size();i>0;i--) {
				para.removeRun(0);
			}
			XWPFRun newInfRun = para.createRun();
			dfs.setDocxFontStyle(newInfRun);
			newInfRun.setText(runText);
		}
	}
	
	private Matcher matcher(String str) {
		Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
		Matcher matcher = pattern.matcher(str);
		return matcher;
	}
	
	Integer fontSize = null;//
	private String matcherReplace(String paraText,Map<String,String> params) {
		
		Matcher matcher;
		String runText = paraText;
		if (this.matcher(paraText).find()) {
			while ((matcher = this.matcher(runText)).find()) {
				String matcherString  = matcher.group(1);
				String replaceString  = params.get(matcherString);
				
				try {
					//判断 占位符 是否包含设定的字体大小数据
					fontSize = Integer.valueOf(matcherString.substring(
									matcherString.toString().length() - 2, 
									matcherString.toString().length()));
					matcherString = matcherString.substring(0, matcherString.toString().length() - 2);
					replaceString  = params.get(matcherString);
				} catch (Exception exp) {
					replaceString  = params.get(matcherString);
				}
				if(replaceString==null) {
					replaceString="";
				}
				runText = matcher.replaceFirst(String.valueOf(replaceString));
			}
		}
		return runText;
	}
	
	private void close(Closeable close) {
		if (close != null) {
			try {
				close.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

}