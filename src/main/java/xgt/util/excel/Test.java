/**
 * 
 */
package xgt.util.excel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;

import xgt.util.excel.model.Pagers;
import xgt.util.excel.templates.DefaultTemplate;
import xgt.util.excel.utils.StyleDecorate;

/**
 * @author Gavin
 *
 */
public class Test {
	public static void main(String[] args) {
		testPagers();
	}
	
	public static void testNormal(){
		String[] header = {"周一","周二","周三","周四","周五","周六","周日"};
		
		List<Object[]> list = new ArrayList<Object[]>();
		for(int i=0; i<6; i++){
			Object[] data = new Object[7];
			data[1] = "String"+i;
			data[2] = new Object();
			data[3] = 10+i;
			data[4] = 20.5+i;
			data[5] = true;
			data[6] = new Date();
			data[0] = null;
			list.add(data);
		}
		
		Template t = TemplateFactory.createTemplate(DefaultTemplate.class, header, list);
		
		CellStyle style = t.createStyle();
		StyleDecorate.decorateAsHeader(style, t.createFont());
		StyleDecorate.decorateBgYellow(style);
		t.setStyle(style, new Region(0,0,0,6));
		
		style = t.createStyle();
		t.setStyle(StyleDecorate.decorateAs￥(style, t.createDataFormat()), new Region(1,6,4,4));
		
		style = t.createStyle();
		t.setStyle(StyleDecorate.decorateAsPercentate(style, t.createDataFormat()), new Region(1,6,3,3));
		
		style = t.createStyle();
		t.setStyle(StyleDecorate.decorateAsDate(style, t.createDataFormat()), new Region(1,6,6,6));
		
		t.setWidthBorder(true,new Region(1,6,0,6));
		
		SheetConfig config = new SheetConfig();
		
		config.addHeightConfig(0, 30f);
		config.addWidthConfig(2, SheetConfig.DEFAULT_WIDTH*3);
		config.addWidthConfig(6, SheetConfig.DEFAULT_WIDTH+3);
		t.setConfig(config);
		
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream("D:/text.xls");
			t.build(fos);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (fos != null) {
					fos.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	
	public static void testPagers(){
		Pagers pagers = new Pagers();
		for(int j=0; j<5; j++){
			String[] header = {"周一","周二","周三","周四","周五","周六","周日"};
			
			List<Object[]> list = new ArrayList<Object[]>();
			for(int i=0; i<6; i++){
				Object[] data = new Object[7];
				data[1] = "String"+i;
				data[2] = new Object();
				data[3] = 10+i;
				data[4] = 20.5+i;
				data[5] = true;
				data[6] = new Date();
				data[0] = null;
				list.add(data);
			}
			MyPager pager = new MyPager("测试"+j, header, list);
			pager.setBorderContents(true);
			pagers.addPager(pager);
		}
		
//		pagers.setTitleOnlyFirstPage(true);
		pagers.setLineSpacing(1);
		
		Template t = TemplateFactory.createTemplate(DefaultTemplate.class, pagers);
		
		SheetConfig config = new SheetConfig();
		config.addHeightConfig(0, 30f);
		config.addWidthConfig(2, SheetConfig.DEFAULT_WIDTH*3);
		config.addWidthConfig(6, SheetConfig.DEFAULT_WIDTH+3);
		t.setConfig(config);
		
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream("D:/text.xls");
			t.build(fos);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (fos != null) {
					fos.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	
	public static void testPager(){
		String[] header = {"周一","周二","周三","周四","周五","周六","周日"};
		
		List<Object[]> list = new ArrayList<Object[]>();
		for(int i=0; i<6; i++){
			Object[] data = new Object[7];
			data[1] = "String"+i;
			data[2] = new Object();
			data[3] = 10+i;
			data[4] = 20.5+i;
			data[5] = true;
			data[6] = new Date();
			data[0] = null;
			list.add(data);
		}
		
		Pager pager = new MyPager("测试", header, list);
		
		Template t = TemplateFactory.createTemplate(DefaultTemplate.class, pager);
		
		CellStyle style = t.createStyle();
		StyleDecorate.decorateAsHeader(style, t.createFont());
		StyleDecorate.decorateBgYellow(style);
		t.setStyle(style, new Region(0,0,0,6));
		
		style = t.createStyle();
		t.setStyle(StyleDecorate.decorateAs￥(style, t.createDataFormat()), new Region(1,6,4,4));
		
		style = t.createStyle();
		t.setStyle(StyleDecorate.decorateAsPercentate(style, t.createDataFormat()), new Region(1,6,3,3));
		
		style = t.createStyle();
		t.setStyle(StyleDecorate.decorateAsDate(style, t.createDataFormat()), new Region(1,6,6,6));
		
		t.setWidthBorder(true,new Region(1,6,0,6));
		
		SheetConfig config = new SheetConfig();
		
		config.addHeightConfig(0, 30f);
		config.addWidthConfig(2, SheetConfig.DEFAULT_WIDTH*3);
		config.addWidthConfig(6, SheetConfig.DEFAULT_WIDTH+3);
		t.setConfig(config);
		
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream("D:/text.xls");
			t.build(fos);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (fos != null) {
					fos.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
}

class MyPager implements Pager{

	/**
	 * @param title
	 * @param headers
	 * @param body
	 */
	public MyPager(String title, String[] headers, List<Object[]> body) {
		super();
		this.title = title;
		this.headers = headers;
		this.body = body;
	}

	private String title;
	
	private String[] headers;
	
	private List<Object[]> body;
	
	private boolean isBorderContents;
	
	
	
	@Override
	public String getTitle() {
		return title;
	}

	@Override
	public int getPageNum() {
		return 0;
	}

	@Override
	public String[] getHeaders() {
		return headers;
	}

	@Override
	public Collection<Object[]> getBody() {
		return body;
	}

	@Override
	public boolean isBorderContents() {
		// TODO Auto-generated method stub
		return this.isBorderContents;
	}

	/**
	 * @param isBorderContents the isBorderContents to set
	 */
	public void setBorderContents(boolean isBorderContents) {
		this.isBorderContents = isBorderContents;
	}
}
