/**
 * 
 */
package last;

import java.util.List;

import xgt.util.easyexcel.last.ExcelCell;
import xgt.util.easyexcel.last.ExcelHandler;
import xgt.util.easyexcel.last.ExcelRow;

/**
 * @author Gavin
 *
 */
public class TestRead {
	public static void main(String[] args) {
		read();
	}
	public static void read(){
		ExcelHandler handler = ExcelHandler.newReadInstance("E:\\test.xlsx");
		List<ExcelRow> rows = handler.read(0, 0);
		for (ExcelRow excelRow : rows) {
			List<ExcelCell> cells = excelRow.getCells();
			for (ExcelCell excelCell : cells) {
				System.out.print(excelCell.getRowNum()+"行-"+excelCell.getColumnNum()+"列:"+excelCell.getContext()+"    ");
			}
			System.out.println();
		}
	}
}
