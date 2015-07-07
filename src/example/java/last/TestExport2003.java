/**
 * 
 */
package last;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Row;

import xgt.util.common.DateUtil;
import xgt.util.easyexcel.last.ExcelCell;
import xgt.util.easyexcel.last.ExcelCellStyle;
import xgt.util.easyexcel.last.ExcelHandler;
import xgt.util.easyexcel.last.ExcelRow;
import xgt.util.easyexcel.last.ExcelSheet;

/**
 * @author Gavin
 *
 */
public class TestExport2003 {
	
	public static void main(String[] args) {
		export2003();
	}
	
	/**
	 * jdk1.6,本地默认环境，可以到处带样式的excel共46300行10列
	 */
	private static void export2003() {
		System.out.println("==start=="+DateUtil.getCurrentDayTime());
		ExcelHandler handler = ExcelHandler.newInstance(ExcelHandler.EXCEL2003);
		ExcelSheet sheet = new ExcelSheet(handler,"测试");
		int column = 10;
		for(int i = 0; i<46300; i++){
			Row row = ExcelRow.createRow(sheet, i);
			for(int j=0; j<column; j++){
				ExcelCell.createCell(row, "测试", j, ExcelCellStyle.newDefaultInstance(handler));
			}
		}
		handler.addSheet(sheet);
		FileOutputStream os = null;
		try {
			os = new FileOutputStream("E:\\test3.xls");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		handler.export(os);
		System.out.println("==end=="+DateUtil.getCurrentDayTime());
	}
}
