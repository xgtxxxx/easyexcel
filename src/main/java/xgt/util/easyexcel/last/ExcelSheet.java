package xgt.util.easyexcel.last;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Sheet;
/**
 * excel sheet类
 * @author xgt
 *
 */
public class ExcelSheet {
	
	private String sheetName;//页签
	
	private short defaultColumnWidth = 12;//当像素3000左右
	
	private short defaultRowHeight = 350;
	
	private Map<Integer,Integer> columnWidths = new HashMap<Integer,Integer>();
	
	private Sheet sheet;
	
	public ExcelSheet(ExcelHandler handler, String sheetName){
		sheet = ExcelUtil.createSheet(handler.getWorkbook(), sheetName);
		sheet.setDefaultColumnWidth(defaultColumnWidth);
		sheet.setDefaultRowHeight(defaultRowHeight);
	}
	/**
	 * 获取sheet名称
	 * @return
	 */
	public String getSheetName(){
		return this.sheetName;
	}
	/**
	 * 设置列宽
	 * @param columnNum
	 * @param width
	 */
	public void setColumnWidth(int columnNum, int width){
		this.columnWidths.put(columnNum, width);
	}
	/**
	 * 获取列宽
	 * @param columnNum 列号
	 * @return
	 */
	public short getColumnWidth(int columnNum){
		if(this.columnWidths.get(columnNum)==null){
			return this.defaultColumnWidth;
		}
		return this.columnWidths.get(columnNum).shortValue();
	}
	/**
	 * 获取列宽
	 * @param columnNum 列号
	 * @return
	 */
	public Map<Integer,Integer> getColumnWidths(){
		return this.columnWidths;
	}
//	/**
//	 * 添加行
//	 * @param row
//	 */
//	public void addRow(ExcelRow row){
//		this.rows.add(row);
//	}
	/**
	 * 设置默认列宽
	 * @param width
	 */
	public void setDefaultColumnWidth(int width){
		defaultColumnWidth = (short)width;
		this.sheet.setDefaultColumnWidth(defaultColumnWidth);
	}
	/**
	 * 获取默认行高
	 * @return
	 */
	public short getDefaultRowHeight() {
		return defaultRowHeight;
	}
	/**
	 * 设置默认行高
	 * @param defaultRowHeight
	 */
	public void setDefaultRowHeight(int defaultRowHeight) {
		this.defaultRowHeight = (short)defaultRowHeight;
		this.sheet.setDefaultRowHeight(getDefaultRowHeight());
	}
	/**
	 * 设置默认行宽
	 * @return
	 */
	public short getDefaultColumnWidth() {
		return defaultColumnWidth;
	}
	public Sheet getSheet() {
		return sheet;
	}
	/**
	 * 在指定行添加分页符
	 * @param row
	 */
	public void addPageBreak(int row){
		sheet.setRowBreak(row);
	}
	/**
	 * 设置自动分页
	 * @param flag
	 */
	public void setAutoBreak(boolean flag){
		sheet.setAutobreaks(flag);
	}
	/**
	 * 设置打印方向
	 * true：横打，false：竖打
	 * @param flag
	 */
	public void setLandscape(boolean flag){
		sheet.getPrintSetup().setLandscape(flag);
	}
	/**
	 * 设置打印边距
	 * @param bottonMargin
	 */
	public void setBottonMargin(double bottonMargin) {
		sheet.setMargin(HSSFSheet.BottomMargin,bottonMargin);
	}
	/**
	 * 设置打印边距
	 */
	public void setRightMargin(double rightMargin) {
		sheet.setMargin(HSSFSheet.RightMargin,rightMargin);
	}
	/**
	 * 设置打印边距
	 */
	public void setLeftMargin(double leftMargin) {
		sheet.setMargin(HSSFSheet.LeftMargin,leftMargin);
	}
	/**
	 * 设置打印边距
	 */
	public void setTopMargin(double topMargin) {
		sheet.setMargin(HSSFSheet.TopMargin,topMargin);
	}
}
