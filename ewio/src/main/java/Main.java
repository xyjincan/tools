
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import cc.watchers.ewio.docxUtils.TemplateWordWriter;
import cc.watchers.ewio.xlsxUtils.ExcelPosition;
import cc.watchers.ewio.xlsxUtils.ExcelWriter;

public class Main {
	
	public static void main(String[] args) {
		

		
		List<List<List<String>>> data = new ArrayList<List<List<String>>>();
		List<List<String>> cdata = new ArrayList<List<String>>();
		List<String> tdata = new ArrayList<String>();
		
		tdata.add("123");tdata.add("123");
		cdata.add(tdata);cdata.add(tdata);cdata.add(tdata);
		data.add(cdata);
		data.add(cdata);
		
		List<List<String>> nwdata = new ArrayList<List<String>>();
		tdata = new ArrayList<String>();
		tdata.add("123");tdata.add("124");tdata.add("125");
		
		nwdata.add(tdata);
		nwdata.add(tdata);
		nwdata.add(tdata);
		
		
		//excel 数据填充例子
		
		ExcelWriter ew = new ExcelWriter("temp.xlsx", "result.xlsx");//
		ew.writeNS(data, new ExcelPosition(0, 1, 0));
		ew.writeNL(nwdata, new ExcelPosition(2, 1,0));
		ew.close();
		
		
		Map<String,String> map = new HashMap<String,String>();
		map.put("title", "word docx 模板替换");
		map.put("你好", "我能吞下玻璃而不伤身体");
		map.put("table", "表格");
		
		
		//word文档例子，只有docx格式支持。
		TemplateWordWriter ww = new TemplateWordWriter();
		ww.TemplateWriter("temp.docx", "result.docx", map);
		
	}

}
