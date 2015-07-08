/**
 * 
 */
package xgt.util.excel;

import java.util.Collection;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.log4j.Logger;

import xgt.util.excel.model.ERow;
import xgt.util.excel.model.Pagers;
import xgt.util.excel.model.RowType;

/**
 * @author Gavin
 *
 */
public final class TemplateFactory {
	
	private static Logger LOG = Logger.getLogger(TemplateFactory.class);
	
	public static Template createTemplate(Class<? extends Template> clazz,String[] header, Collection<Object[]> datas){
		Template t = null;
		try {
			t = clazz.newInstance();
			int index = addHeader(t, header, 0);
			addBody(t, datas, index);
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
				
				index = addHeader(t, pager.getHeaders(), index);
				index = addBody(t, pager.getBody(), index);
				
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
			index = addHeader(t, pager.getHeaders(), index);
			addBody(t, pager.getBody(), index);
		} catch (InstantiationException | IllegalAccessException e) {
			LOG.error(e);
		}
		return t;
	}
	
	private static int addTitle(Template t,String title,int colspan,int index){
		if(StringUtils.isEmpty(title)){
			return index;
		}
		t.addRow(ERow.newTitleInstance(title, index));
		t.merge(new Region(index,index,0,colspan-1));
		return index+1;
	}
	
	private static int addBlankRow(Template t,int index){
		t.addRow(ERow.newBlankInstance(index));
		return index+1;
	}
	
	private static int addBody(Template t, Collection<Object[]> datas, int index){
		for (Object[] objects : datas) {
			t.addRow(ERow.newBodyInstance(objects, index++));
		}
		return index;
	}
	
	private static int addHeader(Template t,String[] header, int index){
		t.addRow(ERow.newInstance(header, index, RowType.HEADER));
		return index+1;
	}
	
}
