/**
 * 
 */
package template;

import java.util.Collection;
import java.util.List;

import xgt.util.excel.Pager;

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