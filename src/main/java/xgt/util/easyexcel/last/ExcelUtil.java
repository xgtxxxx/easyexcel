package xgt.util.easyexcel.last;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {

	@SuppressWarnings("unused")
	private static SimpleDateFormat dateFormat = new SimpleDateFormat(
			"yyyy-MM-dd HH：mm：ss ");
	/**
	 * 字体加粗
	 */
	public static final short FONT_BOLD = Font.BOLDWEIGHT_BOLD;
	/**
	 * 字体不加粗
	 */
	public static final short FONT_BOLD_NOMAL = Font.BOLDWEIGHT_NORMAL;
	/**
	 * 红色字体
	 */
	public static final short FONT_COLOR_RED = Font.COLOR_RED;
	/**
	 * 黑色字体
	 */
	public static final short FONT_COLOR_NORMAL = Font.COLOR_NORMAL;
	/**
	 * 字体加下划线
	 */
	public static final byte UNDERLINE = Font.U_SINGLE_ACCOUNTING;
	/**
	 * 字体不加下划线
	 */
	public static final byte NONE_UNDERLINE = Font.U_NONE;
	/**
	 * 文字居中
	 */
	public static final short ALIGN_CENTER = CellStyle.ALIGN_CENTER;
	/**
	 * 文字左对齐
	 */
	public static final short ALIGN_LEFT = CellStyle.ALIGN_LEFT;
	/**
	 * 文字右对齐
	 */
	public static final short ALIGN_RIGHT = CellStyle.ALIGN_RIGHT;

	public static final short FONT_BIGGER = 16;

	public static final short FONT_BIG = 14;

	public static final short FONT_NORMAL = 12;

	public static final short FONT_SMALL = 10;

	/**
	 * 功能：将Workbook写入Excel文件
	 * 
	 * @param wb
	 *            Workbook
	 * @param absPath
	 *            写入文件的相对路径
	 * @param wbName
	 *            文件名
	 */
//	public static void writeWorkbook(Workbook wb, HttpServletResponse response,String filename) {
//		try {
//			response.reset() ;
//			response.setContentType("application/octet-stream");   
//			response.setHeader("Content-Disposition", "attachment;filename="+new String(filename.getBytes("gb2312"),"ISO-8859-1"));
//			OutputStream os = response.getOutputStream();
//			wb.write(os);
//			os.flush();
//			os.close();
//        } catch (IOException ex) {
//            ex.printStackTrace();
//        }
//
//	}
	/**
	 * 将excel输出到下载流
	 * @param wb
	 * @param fileName
	 */
	public static void writeWorkbook(Workbook wb, OutputStream os) {
		try {
			wb.write(os);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (os != null) {
					os.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	/**
	 * 功能：创建Sheet工作簿
	 * 
	 * @param wb
	 *            Workbook
	 * @param sheetName
	 *            String
	 * @return Sheet
	 */
	public static Sheet createSheet(Workbook wb, String sheetName) {
		Sheet sheet = wb.createSheet(sheetName);
//		sheet.setDefaultColumnWidth(20);
		return sheet;
	}

	/**
	 * 功能：创建Row, default hight
	 * 
	 * @param sheet
	 *            Sheet
	 * @param rowNum
	 *            int
	 * @return Row
	 */
	public static Row createRow(Sheet sheet, int rowNum) {
		Row row = sheet.createRow(rowNum);
		row.setHeight(sheet.getDefaultRowHeight());
		return row;
	}

	/**
	 * 功能：创建CellStyle样式
	 * 
	 * @param wb
	 *            Workbook
	 * @param backgroundColor
	 *            背景色
	 * @param foregroundColor
	 *            前置色
	 * @param font
	 *            字体
	 * @return CellStyle
	 */
	public static CellStyle createCellStyle(Workbook wb, short backgroundColor,
			short foregroundColor, short halign, Font font) {
//		StringBuffer key = new StringBuffer();
		CellStyle cs = wb.createCellStyle();
		cs.setAlignment(halign);
		cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cs.setFillBackgroundColor(backgroundColor);
		cs.setFillForegroundColor(foregroundColor);
		cs.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cs.setFont(font);
		return cs;
	}

	/**
	 * 功能：创建CellStyle样式,自定义属性
	 * 
	 * @param wb
	 *            Workbook
	 * @param font
	 *            字体
	 * @return CellStyle
	 */
	public static CellStyle createCellStyle(Workbook wb, short fontColor, short fontBlod,
			byte fontUnderLine, short fontSize, short fontAlign,
			boolean isBorder) {
		StringBuffer key = new StringBuffer();
		key.append(fontColor).append(fontBlod).append(fontUnderLine).append(fontSize).append(fontAlign).append(isBorder);
		CellStyle cs = wb.createCellStyle();
		Font font = createFont(wb, fontBlod, fontColor, fontSize, fontUnderLine);
		cs.setAlignment(fontAlign);
		cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cs.setFont(font);
		if(isBorder){
			cs.setBorderBottom(HSSFCellStyle.BORDER_THIN);
			cs.setBorderLeft(HSSFCellStyle.BORDER_THIN);
			cs.setBorderRight(HSSFCellStyle.BORDER_THIN);
			cs.setBorderTop(HSSFCellStyle.BORDER_THIN);
		}else{
			cs.setBorderBottom(HSSFCellStyle.BORDER_NONE);
			cs.setBorderLeft(HSSFCellStyle.BORDER_NONE);
			cs.setBorderRight(HSSFCellStyle.BORDER_NONE);
			cs.setBorderTop(HSSFCellStyle.BORDER_NONE);
		}
		cs.setWrapText(true);
		return cs;
	}
	
	/**
	 * 功能：创建CELL
	 * 
	 * @param row
	 *            Row
	 * @param cellNum
	 *            int
	 * @param style
	 *            Style
	 * @return Cell
	 */
	public static Cell createCell(Row row, int cellNum, CellStyle style) {
		Cell cell = row.createCell(cellNum);
		if(style!=null){
			cell.setCellStyle(style);
		}
		return cell;
	}

	/**
	 * 功能：合并单元格
	 * 
	 * @param sheet
	 *            Sheet
	 * @param firstRow
	 *            int
	 * @param lastRow
	 *            int
	 * @param firstColumn
	 *            int
	 * @param lastColumn
	 *            int
	 */
	public static CellRangeAddress mergeCell(Sheet sheet, int firstRow,
			int lastRow, int firstColumn, int lastColumn) {
		CellRangeAddress cra = new ExcelCellRangeAddress(firstRow, lastRow,
				firstColumn, lastColumn);
		sheet.addMergedRegion(cra);
		return cra;
	}

	/**
	 * 功能：创建字体
	 * 
	 * @param wb
	 *            Workbook
	 * @param boldweight
	 *            short
	 * @param color
	 *            short
	 * @return Font
	 */
	public static Font createFont(Workbook wb, short boldweight, short color,
			short size) {
		Font font = wb.createFont();
		font.setBoldweight(boldweight);
		font.setFontName("宋体");
		font.setColor(color);
		font.setFontHeightInPoints(size);
		return font;
	}
	/**
	 * 功能：创建字体
	 * 
	 * @param wb
	 *            Workbook
	 * @param boldweight
	 *            short
	 * @param color
	 *            short
	 * @return Font
	 */
	public static Font createFont(Workbook wb, short boldweight, short color,
			short size, byte underline) {
		Font font = wb.createFont();
		font.setBoldweight(boldweight);
		font.setFontName("宋体");
		font.setColor(color);
		font.setFontHeightInPoints(size);
		font.setUnderline(underline);
		return font;
	}

	/**
	 * 设置合并单元格的边框样式
	 * 
	 * @param sheet
	 *            Sheet
	 * @param ca
	 *            CellRangAddress
	 * @param style
	 *            CellStyle
	 */
	public static void setRegionStyle(Sheet sheet, ExcelCellRangeAddress ca) {
		for (int i = ca.getFirstRow(); i <= ca.getLastRow(); i++) {
			Row row = CellUtil.getRow(i, sheet);
			for (int j = ca.getFirstColumn(); j <= ca.getLastColumn(); j++) {
				Cell cell = CellUtil.getCell(row, j);
				cell.setCellStyle(ca.getCellStyle());
			}
		}
	}
	public static List<ExcelRow> getExcelData(File file, int sheetIndex,int firstRow, int lastRow) {
		return getExcelData(file, sheetIndex, firstRow, lastRow, null);
	}

	/**
	 * 取Excel数据
	 * @return List<String[]>
	 */
	@SuppressWarnings("resource")
	public static List<ExcelRow> getExcelData(File file, int sheetIndex,int firstRow, int lastRow, String ext) {
		List<ExcelRow> rows = new ArrayList<ExcelRow>();
		Workbook wbs = null;
		InputStream io = null;
		try {
			io = new FileInputStream(file);
			
			if (ext != null) {
				if (ext.equalsIgnoreCase("xls")) {
					wbs = new HSSFWorkbook(io);
				}
				
				if (ext.equalsIgnoreCase("xlsx")) {
					wbs = new XSSFWorkbook(io);
				}
			} else {
				if (file.getName().endsWith("xls")) {
					wbs = new HSSFWorkbook(io);
				} else {
					wbs = new XSSFWorkbook(io);
				}
			}
			
			Sheet childSheet = wbs.getSheetAt(sheetIndex);
			
			int first = childSheet.getFirstRowNum();
			int last = childSheet.getLastRowNum();
			if(firstRow>0){
				first = firstRow;
			}
			if(lastRow>0){
				last = lastRow;
			}
			
			for (int i = first; i <= last; i++) {
				Row row = childSheet.getRow(i);
				if (row == null) {
					continue;
				}
				ExcelRow er = ExcelRow.newReadInstance(i);
				int colNum = row.getPhysicalNumberOfCells();
				int j = 0;
				int index = 0;
				while (index < colNum) {
					Cell c = row.getCell(j);
					if(c!=null){
						String s = getStringCellValue(c).trim();
						ExcelCell cell = ExcelCell.newReadInstance(s, j);
						cell.setRowNum(i);
						er.addCell(cell);
						index++;
					}
					j++;
				}
				rows.add(er);
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				io.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return rows;
	}
	/**
	 * 获取第一页的所有数据，包含title
	 * @param file
	 * @return
	 */
	public static List<ExcelRow> getExcelData(File file) {
		return getExcelData(file, 0,-1,-1);
	}
	/**
	 * 获得firstRow以后的数据
	 * @param file
	 * @param firstRow
	 * @return
	 */
	public static List<ExcelRow> getExcelData(File file, int firstRow) {
		return getExcelData(file,0,firstRow,-1);
	}

	private static String getStringCellValue(Cell cell) {
    	if (cell == null) {
            return "";
        }
        String strCell = "";
        switch (cell.getCellType()) {
	        case HSSFCell.CELL_TYPE_STRING:
	            strCell = cell.getStringCellValue();
	            break;
	        case HSSFCell.CELL_TYPE_BOOLEAN:
	            strCell = String.valueOf(cell.getBooleanCellValue());
	            break;
	        case HSSFCell.CELL_TYPE_BLANK:
	            strCell = "";
	            break;
	        case HSSFCell.CELL_TYPE_NUMERIC:
	        	//公式
	        case HSSFCell.CELL_TYPE_FORMULA: {
        		// 判断当前的cell是否为Date
	        	try{
	        		if (HSSFDateUtil.isCellDateFormatted(cell)) {
	        			Date date = cell.getDateCellValue();
	        			SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	        			strCell = sdf.format(date);
	        			
	        			String s = strCell.substring(11);
	        			if("00:00:00".equals(s)){
	        				sdf = new SimpleDateFormat("yyyy-MM-dd");
	        				strCell = sdf.format(date);
	        			}
	        		}else { 
        				double v = cell.getNumericCellValue();
        				long l = (long)v;
        				if(v==l){
        					strCell=String.valueOf(l);
        				}else{
        					strCell = String.valueOf(v);
        				}
	        		}
	        	}catch(Exception e){
	        		e.printStackTrace();
    				strCell = "error";
    			}
        		break;
        	}
	        default:
	            strCell = "";
	            break;
        }
        if (strCell.equals("") || strCell == null) {
            return "";
        }
        return strCell;
    }
}
