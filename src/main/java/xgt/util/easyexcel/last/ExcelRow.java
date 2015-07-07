package xgt.util.easyexcel.last;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;

/**
 * excel行
 * @author xgt
 *
 */
public class ExcelRow {
	//行号
	private int rowNum;
	//行中的cell
	private List<ExcelCell> cells;
	private ExcelRow(){
	}
	/**
	 * 给sheet添加一行数据,供导出用
	 * @param rowNum
	 * @return
	 */
	public static Row createRow(ExcelSheet sheet, int rowNum){
		Row row = ExcelUtil.createRow(sheet.getSheet(), rowNum);
		row.setHeight(sheet.getDefaultRowHeight());
		return row;
	}
	/**
	 * 读取excel实例
	 * @param rowNum
	 * @return
	 */
	public static ExcelRow newReadInstance(int rowNum){
		ExcelRow row = new ExcelRow();
		row.rowNum = rowNum;
		return row;
	}
	/**
	 * 设置行高
	 * @param rowHeight
	 */
	public static void setRowHeight(Row row, int rowHeight) {
		row.setHeight((short)rowHeight);
	}
	/**
	 * 获取行号
	 * @return
	 */
	public int getRowNum() {
		return rowNum;
	}
	/**
	 * 读取Excel时用
	 * 添加一行数据
	 * @param cell
	 */
	public void addCell(ExcelCell cell) {
		if(cells==null){
			cells = new ArrayList<ExcelCell>();
		}
		cells.add(cell);
	}
	/**
	 * 读取excel用
	 * 获取本行数据的所有单元格
	 * @return
	 */
	public List<ExcelCell> getCells() {
		return cells;
	}
	
	/**
	 * 获取本行中指定列的值
	 * @param columnNum 列号
	 * @return
	 */
	public String getContent(int columnNum){
		String context = "";
		for(ExcelCell cell : this.cells){
			if(cell.getColumnNum()==columnNum){
				context = cell.getContext();
			}
		}
		return context;
	}
}
