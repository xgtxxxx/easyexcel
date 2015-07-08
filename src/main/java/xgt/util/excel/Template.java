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

import xgt.util.excel.model.ERow;

/**
 * @author Gavin
 *
 */
public abstract class Template {  
	
	private String name;
	
	private Map<Integer,ERow> rows;
	
	private List<Region> regions;
	
	public void merge(Region region){
		if(regions==null){
			regions = new ArrayList<Region>();
		}
		regions.add(region);
	}
	
	
	public abstract Config getConfig();
	
	public abstract void build(OutputStream os) throws IOException;
	
	
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
