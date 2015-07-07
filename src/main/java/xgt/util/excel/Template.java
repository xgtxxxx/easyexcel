/**
 * 
 */
package xgt.util.excel;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;

import xgt.util.excel.model.ERow;

/**
 * @author Gavin
 *
 */
public abstract class Template {  
	
	private String name;
	
	private Map<Integer,ERow> rows;
	
	private List<Region> regions;
	
	private Map<String,CellStyle> styles = new HashMap<String,CellStyle>();
	
	private Map<String,Boolean> borders;
	
	private boolean borderAll;
	
	private SheetConfig config;
	
	public void merge(Region region){
		if(regions==null){
			regions = new ArrayList<Region>();
		}
		regions.add(region);
	}
	
	public void setStyle(CellStyle style, Region region){
		for(String key : region.getRegionDetails()){
			styles.put(key, style);
		}
	}
	
	public void setStyle(CellStyle style, String key){
		styles.put(key, style);
	}
	
	public CellStyle getStyle(String key){
		return this.styles.get(key);
	}
	
	protected CellStyle getStyleAt(int rowIndex, int columnIndex){
		if(styles==null){
			return getDefaultStyle();
		}
		StringBuffer key = new StringBuffer();
		key.append(rowIndex).append("-").append(columnIndex);
		CellStyle style = styles.get(key.toString());
		return style==null?getDefaultStyle():style;
	}
	
	public abstract void build(OutputStream os) throws IOException;
	
	protected abstract CellStyle getDefaultStyle();
	
	public abstract CellStyle createStyle();
	
	public abstract Font createFont();
	
	public abstract DataFormat createDataFormat();
	
	public void addRow(ERow row){
		if(rows==null){
			rows = new HashMap<Integer,ERow>();
		}
		rows.put(row.getRowIndex(), row);
	}
	
	protected Collection<ERow> getRows(){
		return rows.values();
	}

	/**
	 * @return the regions
	 */
	protected List<Region> getRegions() {
		return regions;
	}

	/**
	 * @return the widthBorder
	 */
	protected boolean isWidthBorder(int rowIndex,int columnIndex) {
		if(borderAll){
			return true;
		}
		if(borders!=null){
			StringBuffer key = new StringBuffer();
			key.append(rowIndex).append("-").append(columnIndex);
			Boolean f = borders.get(key.toString());
			return f==null?false:true;
		}
		return false;
	}

	/**
	 * @param widthBorder the widthBorder to set
	 */
	public void setWidthBorder(boolean widthBorder) {
		this.borderAll = widthBorder;
	}
	
	public void setWidthBorder(boolean widthBoder, Region region){
		if(this.borders==null){
			this.borders = new HashMap<String,Boolean>();
		}
		List<String> keys = region.getRegionDetails();
		for (String key : keys) {
			this.borders.put(key, widthBoder);
		}
	}

	/**
	 * @return the config
	 */
	protected SheetConfig getConfig() {
		return config==null?new SheetConfig():config;
	}

	/**
	 * @param config the config to set
	 */
	public void setConfig(SheetConfig config) {
		this.config = config;
	}

	/**
	 * @return the name
	 */
	public String getName() {
		return name;
	}

	/**
	 * @param name the name to set
	 */
	public void setName(String name) {
		this.name = name;
	}
	
}
