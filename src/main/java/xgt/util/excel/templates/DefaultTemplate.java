/**
 * 
 */
package xgt.util.excel.templates;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import xgt.util.common.DateUtil;
import xgt.util.easyexcel.last.ExcelCellRangeAddress;
import xgt.util.excel.Region;
import xgt.util.excel.SheetConfig;
import xgt.util.excel.Template;
import xgt.util.excel.model.ECell;
import xgt.util.excel.model.ERow;
import xgt.util.excel.utils.StyleDecorate;

/**
 * @author Gavin
 *
 */
public class DefaultTemplate extends Template {
	
	private Workbook wb = new HSSFWorkbook();
	
	private CellStyle defaultStyle = null;
	
	@Override
	protected CellStyle getDefaultStyle() {
		return defaultStyle==null?wb.createCellStyle():defaultStyle;
	}
	
	@Override
	public CellStyle createStyle() {
		return wb.createCellStyle();
	}
	
	@Override
	public Font createFont() {
		return wb.createFont();
	}
	
	@Override
	public DataFormat createDataFormat() {
		return wb.createDataFormat();
	}
	
	@Override
	public void build(OutputStream os) throws IOException {
		Sheet sheet = wb.createSheet(getName());
		initSheetWidthConfig(sheet);
		buildRows(sheet);
		setMerge(sheet);
		wb.write(os);
	}
	
	private void initSheetWidthConfig(Sheet sheet){
		SheetConfig config = this.getConfig();
		for (int index : config.getKeysOfWidths()) {
			sheet.setColumnWidth(index, config.getWidth(index)*256);
		}
	}
	
	private void buildRows(Sheet sheet){
		SheetConfig config = this.getConfig();
		for (ERow er : getRows()) {
			Row row = sheet.createRow(er.getRowIndex());
			if(er.getRowHeight()>0f){
				row.setHeightInPoints(er.getRowHeight());
			}else{
				row.setHeightInPoints(config.getHeight(er.getRowIndex()));
			}
			buildCells(row,er);
		}
	}
	
	private void buildCells(Row row,ERow er){
		for (ECell ec : er.getCells()) {
			Cell cell = row.createCell(ec.getColumnIndex());
			Object v = ec.getValue();
			if(v instanceof Boolean){
				cell.setCellValue((Boolean)v);
			}else if(v instanceof Number){
				cell.setCellValue(((Number)v).doubleValue());
			}else if(v instanceof Date){
				cell.setCellValue(DateUtil.formatDate((Date)v));
			}else if(v instanceof String){
				cell.setCellValue((String)v);
			}else if(v instanceof RichTextString){
				cell.setCellValue((RichTextString)v);
			}else{
				cell.setCellValue(String.valueOf(v));
			}
			
			CellStyle style = getStyleAt(cell.getRowIndex(), cell.getColumnIndex());
			if(isWidthBorder(ec.getRowIndex(),ec.getColumnIndex())){
				StyleDecorate.addBorder(style);
			}
			cell.setCellStyle(style);
		}
	}
	
	private void setMerge(Sheet sheet){
		List<Region> regions = getRegions();
		if(regions==null){
			return;
		}
		for (Region region : regions) {
			CellRangeAddress cra = new ExcelCellRangeAddress(region.getStartRow(),region.getEndRow(),region.getStartColumn(),region.getEndColumn());
			sheet.addMergedRegion(cra);
		}
	}

}
