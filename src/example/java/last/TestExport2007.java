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
public class TestExport2007 {
	public static void main(String[] args) {
		export2007();
	}
	/**
	 * jdk1.6,本地默认环境，可以到处带样式的excel共
	 * 用时22秒到处50万行数据没有内存溢出
	 */
	private static void export2007() {
		System.out.println("==start=="+DateUtil.getCurrentDayTime());
		ExcelHandler handler = ExcelHandler.newInstance(ExcelHandler.EXCEL2007);
		ExcelSheet sheet = new ExcelSheet(handler,"测试");
		int column = 10;
		for(int i = 0; i<500000; i++){
			Row row = ExcelRow.createRow(sheet, i);
			for(int j=0; j<column; j++){
				ExcelCell.createCell(row, "测试", j, ExcelCellStyle.newDefaultInstance(handler));
			}
		}
		handler.addSheet(sheet);
		FileOutputStream os = null;
		try {
			os = new FileOutputStream("D:\\test.xlsx");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		handler.export(os);
		System.out.println("==end=="+DateUtil.getCurrentDayTime());
	}

}
