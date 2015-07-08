/**
 * 
 */
package xgt.util.excel.model;

/**
 * @author Gavin
 *
 */
public enum RowType {
	TITLE("title"), 
	HEADER("header"),
	BODY("body"),
	FOOTER("footer"), 
	BLANK("blank"), 
	DEFAULT("default");
	
	private String name;

	private RowType(String name) {
		this.name = name;
	}

	/**
	 * @return the name
	 */
	public String getName() {
		return name;
	}
}