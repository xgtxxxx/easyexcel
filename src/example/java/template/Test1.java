/**
 * 
 */
package template;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;

import xgt.util.excel.Config;
import xgt.util.excel.Region;
import xgt.util.excel.Template;
import xgt.util.excel.TemplateFactory;
import xgt.util.excel.templates.DefaultTemplate;
import xgt.util.excel.utils.StyleDecorate;

/**
 * @author Gavin
 *
 */
public class Test1 {
	public static void main(String[] args) {
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
		Config config = t.getConfig();
		CellStyle style = config.createStyle();
		StyleDecorate.decorateAsHeader(style, config.createFont());
		StyleDecorate.decorateBgYellow(style);
		config.setStyle(style, new Region(0,0,0,6));
		
		style = config.createStyle();
		config.setStyle(StyleDecorate.decorateAs￥(style, config.createDataFormat()), new Region(1,6,4,4));
		
		style = config.createStyle();
		config.setStyle(StyleDecorate.decorateAsPercentate(style, config.createDataFormat()), new Region(1,6,3,3));
		
		style = config.createStyle();
		config.setStyle(StyleDecorate.decorateAsDate(style, config.createDataFormat()), new Region(1,6,6,6));
		
		config.addRowHeight(0, 30f);
		config.addColumnWidth(2, Config.DEFAULT_WIDTH*3);
		config.addColumnWidth(6, Config.DEFAULT_WIDTH+3);
		
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream("D:/text.xlsx");
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
