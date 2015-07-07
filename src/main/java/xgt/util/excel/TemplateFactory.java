/**
 * 
 */
package xgt.util.excel;

import java.util.Collection;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.CellStyle;

import xgt.util.excel.model.ERow;
import xgt.util.excel.model.ERow.RowType;
import xgt.util.excel.model.Pagers;
import xgt.util.excel.utils.StyleDecorate;

/**
 * @author Gavin
 *
 */
public final class TemplateFactory {
	
	private static final String KEY_TITLE = "title";
	
	private static final String KEY_HEADER = "header";
	
	private static final String KEY_BODY = "body";
	
	private static Logger LOG = Logger.getLogger(TemplateFactory.class);
	
	public static Template createTemplate(Class<? extends Template> clazz,String[] header, Collection<Object[]> datas){
		Template t = null;
		try {
			t = clazz.newInstance();
			int index = addHeader(t, header, 0, true);
			addBody(t, datas, index, true);
		} catch (InstantiationException | IllegalAccessException e) {
			LOG.error(e);
		}
		return t;
	}
	
	public static Template createTemplate(Class<? extends Template> clazz, Pagers pagers){
		Template t = null;
		try {
			t = clazz.newInstance();
			int index = 0;
			List<Pager> ps = pagers.getPagers();
			for (int i = 0; i<ps.size(); i++) {
				Pager pager = ps.get(i);
				if(pagers.isTitleOnlyFirstPage()){
					if(i==0){
						index = addTitle(t,pager.getTitle(),pager.getHeaders().length,index++);
					}
				}else{
					index = addTitle(t,pager.getTitle(),pager.getHeaders().length,index++);
				}
				
				index = addHeader(t, pager.getHeaders(), index, pager.isBorderContents());
				index = addBody(t, pager.getBody(), index, pager.isBorderContents());
				
				for(int j=0; j<pagers.getLineSpacing(); j++){
					index = addBlankRow(t, index);
				}
			}
			
		} catch (InstantiationException | IllegalAccessException e) {
			LOG.error(e);
		}
		return t;
	}
	
	public static Template createTemplate(Class<? extends Template> clazz, Pager pager){
		Template t = null;
		try {
			t = clazz.newInstance();
			int index = 0;
			if(StringUtils.isNotEmpty(pager.getTitle())){
				index = addTitle(t,pager.getTitle(),pager.getHeaders().length,index);
			}
			index = addHeader(t, pager.getHeaders(), index, pager.isBorderContents());
			addBody(t, pager.getBody(), index, pager.isBorderContents());
		} catch (InstantiationException | IllegalAccessException e) {
			LOG.error(e);
		}
		return t;
	}
	
	private static int addTitle(Template t,String title,int colspan,int index){
		if(StringUtils.isEmpty(title)){
			return index;
		}
		t.addRow(ERow.newInstance(title, index));
		t.merge(new Region(index,index,0,colspan-1));
		
		CellStyle style = t.getStyle(KEY_TITLE);
		if(style==null){
			style = t.createStyle();
			StyleDecorate.decorateAsTitle(style, t.createFont());
			t.setStyle(style, KEY_TITLE);
		}
		t.setStyle(style, new Region(index,index,0,0));
		
		return index+1;
	}
	
	private static int addBlankRow(Template t,int index){
		t.addRow(ERow.newInstance(index));
		return index+1;
	}
	
	private static int addBody(Template t, Collection<Object[]> datas, int index, boolean isBorder){
		int start = index;
		for (Object[] objects : datas) {
			t.addRow(ERow.newInstance(objects, index++));
		}
		
		if(isBorder){
			CellStyle style = t.getStyle(KEY_BODY);
			if(style==null){
				style = t.createStyle();
				StyleDecorate.addBorder(style);
				t.setStyle(style, KEY_BODY);
			}
			t.setStyle(style, new Region(start,index,0,datas.iterator().next().length-1));
		}
		return index;
	}
	
	private static int addHeader(Template t,String[] header, int index, boolean isBorder){
		t.addRow(ERow.newInstance(header, index, RowType.HEADER));
		CellStyle style = t.getStyle(KEY_HEADER);
		if(style==null){
			style = t.createStyle();
			StyleDecorate.decorateAsHeader(style, t.createFont());
			if(isBorder){
				StyleDecorate.addBorder(style);
			}
			t.setStyle(style, KEY_HEADER);
		}
		t.setStyle(style, new Region(index,index,0,header.length-1));
		return index+1;
	}
	
}
