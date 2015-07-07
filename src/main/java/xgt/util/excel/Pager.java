/**
 * 
 */
package xgt.util.excel;

import java.util.Collection;

/**
 * @author Gavin
 *
 */
public interface Pager {
	public String getTitle();
	
	public int getPageNum();
	
	public String[] getHeaders();
	
	public Collection<Object[]> getBody();
	
	public boolean isBorderContents();
}
