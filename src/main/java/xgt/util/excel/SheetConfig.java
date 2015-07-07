/**
 * 
 */
package xgt.util.excel;

import java.util.Collection;
import java.util.HashMap;
import java.util.Map;

/**
 * @author Gavin
 *
 */
public class SheetConfig {
	
	public static final int DEFAULT_WIDTH = 8;
	
	public static final float DEFAULT_HEIGHT = 12.75f;
	
	private int defaultWidth = DEFAULT_WIDTH;
	
	private float defaultHeight = DEFAULT_HEIGHT;
	
	private Map<Integer,Integer> widths = new HashMap<Integer,Integer>();
	
	private Map<Integer,Float> heights = new HashMap<Integer,Float>();
	
	public int getWidth(int index){
		Integer w = widths.get(index);
		return w==null?defaultWidth:w;
	}
	
	public float getHeight(int index){
		Float h = heights.get(index);
		return h==null?defaultHeight:h;
	}
	
	public void addWidthConfig(int index,int width){
		widths.put(index, width);
	}
	
	public void addHeightConfig(int index, float height){
		heights.put(index, height);
	}
	
	/**
	 * @return the defaultWidth
	 */
	public int getDefaultWidth() {
		return defaultWidth;
	}

	/**
	 * @param defaultWidth the defaultWidth to set
	 */
	public void setDefaultWidth(int defaultWidth) {
		this.defaultWidth = defaultWidth;
	}

	/**
	 * @return the defaultHeight
	 */
	public float getDefaultHeight() {
		return defaultHeight;
	}

	/**
	 * @param defaultHeight the defaultHeight to set
	 */
	public void setDefaultHeight(float defaultHeight) {
		this.defaultHeight = defaultHeight;
	}

	/**
	 * @return the widths
	 */
	public Collection<Integer> getKeysOfWidths() {
		return widths.keySet();
	}

	/**
	 * @return the heights
	 */
	public Collection<Integer> getKeysOfHeights() {
		return heights.keySet();
	}
}
