package xgt.util.easyexcel.last;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;


/**
 * excel单元格
 * @author xgt
 *
 */
public class ExcelCell {
	//内容
	private String context;
	//行号
	private int rowNum;
	//列号
	private int columnNum;
	/**
	 * 
	 * @param content
	 */
	private ExcelCell(){
	}
	/**
	 * 合并单元格
	 * @param startRow
	 * @param endRow
	 * @param startColumn
	 * @param endColumn
	 */
	public static void setMerge(Cell cell, int startRow, int endRow, int startColumn, int endColumn){
		if((endRow-startRow>=0)&&(endColumn-startColumn>=0)){
			ExcelCellRangeAddress ca = (ExcelCellRangeAddress)ExcelUtil.mergeCell(cell.getSheet(), startRow, endRow, startColumn, endColumn);
			ca.setCellStyle(cell.getCellStyle());
			ExcelCellRangeAddress.addRanges(ca);
		}
	}
	/**
	 * 创建单元格
	 * @param content 内容
	 * @param column 列号
	 */
	public static Cell createCell(Row row, String content, int column){
		Cell cell = ExcelUtil.createCell(row, column, null);
		cell.setCellValue(content);
		return cell;
	}
	/**
	 * 创建单元格,有边框，居中
	 * @param content 内容
	 * @param column 列号
	 */
	public static Cell createDefaultCell(Row row, String content, int column){
		Cell cell = ExcelUtil.createCell(row, column, ExcelCellStyle.newDefaultInstance(row.getSheet().getWorkbook()));
		cell.setCellValue(content);
		return cell;
	}
	/**
	 * 创建单元格
	 * @param content 内容
	 * @param column 列号
	 * @param style 单元格样式
	 */
	public static Cell createCell(Row row, String content, int column, CellStyle style){
		Cell cell = ExcelUtil.createCell(row, column, style);
		cell.setCellValue(content);
		return cell;
	}
	
	/**
	 * 读取excel实例
	 * @param string
	 * @param j
	 * @return
	 */
	public static ExcelCell newReadInstance(String string, int j) {
		ExcelCell cell = new ExcelCell();
		cell.columnNum = j;
		cell.context = string;
		return cell;
	}
	
	public String getContext() {
		return context;
	}
	public int getRowNum() {
		return rowNum;
	}
	public void setRowNum(int rowNum) {
		this.rowNum = rowNum;
	}
	public int getColumnNum() {
		return columnNum;
	}
	public void setColumnNum(int columnNum) {
		this.columnNum = columnNum;
	}
	public void setContext(String context) {
		this.context = context;
	}
}
