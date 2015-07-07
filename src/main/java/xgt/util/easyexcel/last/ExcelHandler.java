package xgt.util.easyexcel.last;

import java.io.File;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 * excel处理类
 * @author xgt
 *
 */
public class ExcelHandler {
	
	public static final String EXCEL2007 = "xlsx";
	
	public static final String EXCEL2003 = "xls";
	
	private Workbook wb = null;
	
	private File file;
	
	/** 导出excle的文件名 */
	private String fileName;
	
	private String path;
	
//	private HttpServletResponse response;

	/** 存放要显示的Sheet */
	private List<ExcelSheet> sheets = new ArrayList<ExcelSheet>();
	
	private ExcelHandler(){
		
	}
	/**
	 * 导出实例
	 * @return
	 */
	public static ExcelHandler newInstance(String fileType){
		ExcelHandler handler = new ExcelHandler();
		if(fileType.equals(EXCEL2003)){
			handler.wb = new HSSFWorkbook();
		}else if(fileType.equals(EXCEL2007)){
//				handler.wb = new XSSFWorkbook();
			handler.wb = new SXSSFWorkbook(500);
		}
		return handler;
	}
	/**
	 * 读取excel实例
	 * @param path excel路径
	 * @return
	 */
	public static ExcelHandler newReadInstance(String path){
		ExcelHandler handler = new ExcelHandler();
		handler.setPath(path);
		int index = path.lastIndexOf(".")+1;
		String fileType = path.substring(index);
		if(fileType!=null){
			if(fileType.equals("xls")){
				handler.wb = new HSSFWorkbook();
			}else{
				handler.wb = new XSSFWorkbook();
//				handler.wb = new SXSSFWorkbook(500);
			}
		}
		return handler;
	}
	/**
	 * 读取excel实例
	 * @param path excel路径
	 * @return
	 */
	public static ExcelHandler newReadInstance(File file,String fileName){
		ExcelHandler handler = new ExcelHandler();
		handler.file = file;
		handler.fileName = fileName;
		if(fileName.endsWith("xls")){
			handler.wb = new HSSFWorkbook();
		}else{
			handler.wb = new XSSFWorkbook();
		}
		return handler;
	}
	/**
	 * 添加sheet
	 * @param sheet
	 */
	public void addSheet(ExcelSheet sheet){
		this.sheets.add(sheet);
	}
	
	private void initColumnWinth(ExcelSheet es){
		Map<Integer,Integer> columns = es.getColumnWidths();
		Set<Integer> keys = columns.keySet();
		for (Integer key : keys) {
			int width = columns.get(key);
			if(width!=0){
				es.getSheet().setColumnWidth(key, width*100);
			}else{
				es.getSheet().setColumnWidth(key, es.getDefaultColumnWidth()*100);
			}
		}
	} 
	/**
	 * 读取excel
	 * @param sheet
	 * @param startRow
	 * @param endRow
	 * @return
	 */
	public List<ExcelRow> read(int sheet, int startRow, int endRow){
//		File file = new File(this.path);
		if(file==null){
			file = new File(path);
		}
		List<ExcelRow> rows = null;
		try {
//			List<List<String>> list = null;
			if (this.fileName != null) {
				String ext = this.fileName.substring(this.fileName.lastIndexOf(".") + 1);
				rows = ExcelUtil.getExcelData(file, sheet, startRow, endRow, ext);
			} else {
				rows = ExcelUtil.getExcelData(file, sheet, startRow, endRow);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return rows;
	}
	/**
	 * 读取excel
	 * @param sheet
	 * @param startRow
	 * @return
	 */
	public List<ExcelRow> read(int sheet, int startRow){
		return this.read(sheet, startRow, -1);
	}
	
	/**
	 * 导出excel
	 */
	public void export(OutputStream os){
		for(int i = 0; i<this.sheets.size(); i++){
			ExcelSheet table = this.sheets.get(i);
			String sheetName = table.getSheetName();
			if(sheetName==null || sheetName.length()<=0){
				sheetName = "sheet".concat(String.valueOf(i));
			}
			Sheet sheet = table.getSheet();
			this.initColumnWinth(table);
			
			List<ExcelCellRangeAddress> ranges = ExcelCellRangeAddress.getRanges();
			if(ranges.size()>0){
				for(ExcelCellRangeAddress ca : ranges){
					ExcelUtil.setRegionStyle(sheet, ca);
				}
			}
		}
		ExcelUtil.writeWorkbook(wb, os);
		this.clear();
	}
	
	public void clear(){
		ExcelCellRangeAddress.clear();
		ExcelCellStyle.clear();
	}
	
	public String getFileName() {
		return fileName;
	}

	public void setFileName(String fileName) {
		this.fileName = fileName;
	}

	public String getPath() {
		return path;
	}

	public void setPath(String path) {
		this.path = path;
	}

	public List<ExcelSheet> getSheets() {
		return sheets;
	}

	public void setSheets(List<ExcelSheet> sheets) {
		this.sheets = sheets;
	}
//	public HttpServletResponse getResponse() {
//		return response;
//	}
//	public void setResponse(HttpServletResponse response) {
//		this.response = response;
//	}
	public Workbook getWorkbook() {
		return wb;
	}
}
