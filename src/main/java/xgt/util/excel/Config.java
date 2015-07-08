/**
 * 
 */
package xgt.util.excel;

import java.util.Collection;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;

import xgt.util.excel.model.RowType;

/**
 * @author Gavin
 *
 */
public abstract class Config {
	
	protected Config(){
		init();
	}
	
	public static final int DEFAULT_WIDTH = 8;
	
	public static final float DEFAULT_HEIGHT = 12.75f;
	
	private Map<String,CellStyle> styles = new HashMap<String,CellStyle>();
	
	private Map<Integer,Integer> widths = new HashMap<Integer,Integer>();
	
	private Map<Integer,Float> heigths= new HashMap<Integer,Float>();
	
	private Map<RowType,Boolean> borders = new HashMap<RowType,Boolean>();
	
	private int defaultWidth = DEFAULT_WIDTH;
	
	private float defaultHeight = DEFAULT_HEIGHT;
	
	public boolean isBorder(RowType key){
		Boolean f = borders.get(key);
		return f==null?false:f;
	}
	
	public void addBorder(RowType key){
		borders.put(key, true);
	}
	
	public int getColumnWidth(int index){
		Integer w = widths.get(index);
		return w==null?defaultWidth:w;
	}
	
	public float getRowHeight(int index){
		Float w = heigths.get(index);
		return w==null?defaultHeight:w;
	}
	
	public void addColumnWidth(int index,int width){
		this.widths.put(index, width);
	}
	
	public void addRowHeight(int index, float height){
		this.heigths.put(index, height);
	}
	
	public CellStyle getStyle(String key){
		CellStyle style = this.styles.get(key);
		return style==null?createStyle():style;
	}
	
	public CellStyle getStyle(RowType key){
		return getStyle(key.getName());
	}
	
	public CellStyle getStyleAt(int rowIndex, int columnIndex){
		StringBuffer key = new StringBuffer();
		key.append(rowIndex).append("-").append(columnIndex);
		CellStyle style = styles.get(key.toString());
		return style==null?getDefaultStyle():style;
	}
	
	public Collection<Integer> getKeysOfWidth(){
		return this.widths.keySet();
	}
	
	public Collection<Integer> getKeysOfHeight(){
		return this.heigths.keySet();
	}
	
	private CellStyle getDefaultStyle(){
		CellStyle s = this.styles.get(RowType.DEFAULT.getName());
		return s==null?this.createStyle():s;
	}
	
	public void setStyle(CellStyle style, Region region){
		for(String key : region.getRegionDetails()){
			styles.put(key, style);
		}
	}
	
	public void setStyle(CellStyle style, RowType key){
		styles.put(key.getName(), style);
	}
	
	public abstract CellStyle createStyle();
	
	public abstract Font createFont();
	
	public abstract DataFormat createDataFormat();
	
	protected abstract void init();
	
	/**
	 * @return the defaultWidth
	 */
	public int getDefaultWidth() {
		return defaultWidth;
	}

	/**
	 * @return the defaultHeight
	 */
	public float getDefaultHeight() {
		return defaultHeight;
	}

	/**
	 * @param defaultWidth the defaultWidth to set
	 */
	public void setDefaultWidth(int defaultWidth) {
		this.defaultWidth = defaultWidth;
	}

	/**
	 * @param defaultHeight the defaultHeight to set
	 */
	public void setDefaultHeight(float defaultHeight) {
		this.defaultHeight = defaultHeight;
	}
}
