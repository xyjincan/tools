package cc.watchers.ewio.docxUtils;


import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.VerticalAlign;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * 用于保存删除的 XWPFRun属性
 */
public class DocxFontStyle {
	
	public int DefaultFontSize = 16;//默认字体大小。
	
	String fontFamily;
	String color;
	boolean bold;
	boolean italic;
	UnderlinePatterns underlinePatterns;
	int possition;
	VerticalAlign verticalAlign;
	int kerning;
	boolean capitalized;
	boolean doubleStrikethrough;
	boolean embossed;
	boolean imprinted;
	boolean shadow;
	boolean mallCaps;
	UnderlinePatterns underline;
	
	
	Integer fontSize = null;//

	public DocxFontStyle(XWPFRun run,Integer fontSize) {
		
		this.fontSize=fontSize;
		fontFamily = run.getFontFamily();
		color = run.getColor();
		bold = run.isBold();
		italic = run.isItalic();
		underlinePatterns = run.getUnderline();
		possition = run.getTextPosition();
		verticalAlign = run.getSubscript();
		kerning = run.getKerning();
		capitalized = run.isCapitalized();
		doubleStrikethrough = run.isDoubleStrikeThrough();
		embossed = run.isEmbossed();
		imprinted = run.isImprinted();
		shadow = run.isShadowed();
		mallCaps = run.isSmallCaps();
		underline = run.getUnderline();
	}

	public void setDocxFontStyle(XWPFRun newRun) {
		
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
		
		if(fontSize!=null) {
			newRun.setFontSize(fontSize);
		}else {
			newRun.setFontSize(DefaultFontSize);
		}
	}
}
