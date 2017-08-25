package cc.watchers.ewio.xlsxUtils;

/**
 * 
 * 写文件定位
 *
 */
public class ExcelPosition {
	
	public int s;//文件第几页0开始
	public int l;//页内第几行0开始
	public int c;//行内第几格0开始
	
	public ExcelPosition(int s,int l,int c) {
		this.s=s;
		this.l=l;
		this.c=c;
	}
	

}
