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

import xgt.util.excel.Config;
import xgt.util.excel.Template;
import xgt.util.excel.TemplateFactory;
import xgt.util.excel.model.Pagers;
import xgt.util.excel.templates.DefaultTemplate;

/**
 * @author Gavin
 *
 */
public class Test3 {
	public static void main(String[] args) {
		Pagers pagers = new Pagers();
		for (int j = 0; j < 5; j++) {
			String[] header = { "周一", "周二", "周三", "周四", "周五", "周六", "周日" };

			List<Object[]> list = new ArrayList<Object[]>();
			for (int i = 0; i < 6; i++) {
				Object[] data = new Object[7];
				data[1] = "String" + i;
				data[2] = new Object();
				data[3] = 10 + i;
				data[4] = 20.5 + i;
				data[5] = true;
				data[6] = new Date();
				data[0] = null;
				list.add(data);
			}
			MyPager pager = new MyPager("测试" + j, header, list);
			pagers.addPager(pager);
		}

		pagers.setTitleOnlyFirstPage(true);
		pagers.setLineSpacing(1);

		Template t = TemplateFactory.createTemplate(DefaultTemplate.class,
				pagers);
		
		Config config = t.getConfig();
		
		config.setDefaultWidth(config.getDefaultWidth()*2);

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
