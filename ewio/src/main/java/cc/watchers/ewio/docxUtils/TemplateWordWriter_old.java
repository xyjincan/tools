package cc.nitsc.ewio.docxUtils;

import java.io.Closeable;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.VerticalAlign;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import cc.nitsc.logstatistics.Utils;

/**
 * 
 * 只支持了docx格式
 * 占位符${xxxxss}
 * xxxx标识文档中需要替换的定位，ss表示该位置需要的字体大小，默认16
 */

public class TemplateWordWriter {

	public int DefaultFontSize = 16;
	
	public XWPFDocument doc = null;
	public String resultFile;
	
	
	
	Integer fontSize = null;//
	
	public void TemplateWriter(String template,String resultFilePath,Map<String, String> params){

		try {
			InputStream is = new FileInputStream(template);
			doc = new XWPFDocument(is);
			resultFile = resultFilePath;
			// 替换文档里面的变量
			this.replaceInPara(doc, params);
			// 替换表格里面的变量
			this.replaceInTable(doc, params);
			
		} catch (IOException e) {
			new IOException(template+" 写入 "+resultFilePath+" 失败").printStackTrace();
		}
	}
	
	
	private void replaceInPara(XWPFDocument docx, Map<String, String> params) {
		Iterator<?> iterator = docx.getParagraphsIterator();
		XWPFParagraph para;
		while (iterator.hasNext()) {
			fontSize = null;
			para = (XWPFParagraph) iterator.next();
			this.replaceInPara(para, params);
		}
	}

	private void replaceInPara(XWPFParagraph para, Map<String, String> params) {

		List<XWPFRun> runs;
		Matcher matcher;
		if (this.matcher(para.getParagraphText()).find()) {
			runs = para.getRuns();
			for (int i = 0; i < runs.size(); i++) {
				XWPFRun run = runs.get(i);
				String runText = run.toString();
				//System.out.println("Template:"+runText);
				matcher = this.matcher(runText);
				if (matcher.find()) {
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
					// 直接调用XWPFRun的setText()方法设置文本时，在底层会重新创建一个XWPFRun，把文本附加在当前文本后面，
					// 所以我们不能直接设值，需要先删除当前run,然后再自己手动插入一个新的run。

					String fontFamily = run.getFontFamily();
					String color = run.getColor();
					boolean bold = run.isBold();
					boolean italic = run.isItalic();
					UnderlinePatterns underlinePatterns = run.getUnderline();
					int possition = run.getTextPosition();
					VerticalAlign verticalAlign = run.getSubscript();
					int kerning = run.getKerning();
					boolean capitalized = run.isCapitalized();
					boolean doubleStrikethrough = run.isDoubleStrikeThrough();
					boolean embossed = run.isEmbossed();
					boolean imprinted = run.isImprinted();
					boolean shadow = run.isShadowed();
					boolean mallCaps = run.isSmallCaps();
					UnderlinePatterns underline = run.getUnderline();

					//System.out.println(run.toString()+" turn to :"+runText);
					
					//fontSize = null;
					try {
						fontSize = Integer.valueOf(
								run.toString().substring(run.toString().length() - 3, run.toString().length() - 1));
						
					} catch (Exception exp) {
					}
					
					if(fontSize!=null) {
						System.out.print("FontSize: "+fontSize+"。");	
					}
					System.out.println(run.toString()+" turn to :"+runText);
					
					// 删除原有文段，再追加新段
					para.removeRun(i);
					XWPFRun newRun = para.insertNewRun(i);

					newRun.setText(runText);
					if (fontSize != null) {
						newRun.setFontSize(fontSize);// 字体大小
					} else {
						newRun.setFontSize(DefaultFontSize);
					}
					newRun.setFontFamily(fontFamily);
					newRun.setColor(color);
					newRun.setBold(bold);
					newRun.setUnderline(underlinePatterns);
					newRun.setTextPosition(possition);
					newRun.setSubscript(verticalAlign);
					newRun.setKerning(kerning);
					newRun.setItalic(italic);
					newRun.setCapitalized(capitalized);
					// newRun.setCharacterSpacing(characterSpacing);
					newRun.setDoubleStrikethrough(doubleStrikethrough);
					newRun.setEmbossed(embossed);
					newRun.setImprinted(imprinted);
					newRun.setShadow(shadow);
					newRun.setSmallCaps(mallCaps);
					// newRun.setSubscript(subscript);
					newRun.setUnderline(underline);

				}
			}
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
						this.replaceInPara(para, params);
					}
				}
			}
		}
	}

	private Matcher matcher(String str) {
		Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
		Matcher matcher = pattern.matcher(str);
		return matcher;
	}


	
	private String matcherReplace(String paraText,Map<String,String> params) {
		
		Matcher matcher;
		String runText = paraText;
		if (this.matcher(paraText).find()) {
			fontSize = null;
			//反复
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
	
	public void closeDoc() {
		
		OutputStream os;
		try {
			os = new FileOutputStream(resultFile);
			doc.write(os);
			close(os);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}