/**
 * 
 */
package last;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import xgt.util.easyexcel.last.ExcelCell;
import xgt.util.easyexcel.last.ExcelCellStyle;
import xgt.util.easyexcel.last.ExcelHandler;
import xgt.util.easyexcel.last.ExcelRow;
import xgt.util.easyexcel.last.ExcelSheet;

/**
 * @author Gavin
 *
 */
public class TestExportReport {
	public static void main(String[] args) {
		export();
	}
	public static void export(){
		ExcelHandler handler = ExcelHandler.newInstance(ExcelHandler.EXCEL2007);
		ExcelSheet sheet = new ExcelSheet(handler,"测试");
		sheet.setColumnWidth(0, 10);
		sheet.setColumnWidth(1, 10);
		sheet.setColumnWidth(2, 15);
		sheet.setColumnWidth(3, 15);
		sheet.setColumnWidth(14, 30);
		
		//打印设置
		sheet.setAutoBreak(false);
		sheet.setLandscape(true);
		sheet.setLeftMargin(0);
//		sheet.setDefaultColumnWidth(12);
//		sheet.setDefaultRowHeight(500);
		
		Row row = ExcelRow.createRow(sheet, 0);
		Cell cell = ExcelCell.createCell(row, "某单位2013年06月基本医疗保险基金拨付凭证", 0, ExcelCellStyle.newInstance(handler, ExcelCellStyle.BORDER_NONE_CENTER));
//		ExcelCell cell = ExcelCell.createCell("某单位2013年06月基本医疗保险基金拨付凭证",0,ExcelCellStyle.newInstance(ExcelCellStyle.BORDER_NONE_CENTER));
		ExcelCell.setMerge(cell, 0, 0, 0, 18);
		
		row = ExcelRow.createRow(sheet, 1);
		cell = ExcelCell.createCell(row,"定点医疗机构名称：某单位医院", 0,ExcelCellStyle.newInstance(handler,ExcelCellStyle.BORDER_NONE_LEFT));
		ExcelCell.setMerge(cell, 1, 1, 0, 6);
//		//row.addCell(cell);
		cell = ExcelCell.createCell(row,"付款所属期：2011年06月", 7, ExcelCellStyle.newInstance(handler,ExcelCellStyle.BORDER_NONE_RIGHT));
		ExcelCell.setMerge(cell, 1, 1, 7, 18);
//		//row.addCell(cell);
//		//table.addRow(row);
		
		row = ExcelRow.createRow(sheet,2);
		cell = ExcelCell.createCell(row, "定点医疗机构编码：1207", 0, ExcelCellStyle.newInstance(handler, ExcelCellStyle.BORDER_NONE_LEFT));
		ExcelCell.setMerge(cell, 2, 2, 0, 8);
//		//row.addCell(cell);
		cell = ExcelCell.createCell(row, "财务支付流水号：2011060288",9, ExcelCellStyle.newInstance(handler, ExcelCellStyle.BORDER_NONE_LEFT));
		ExcelCell.setMerge(cell, 2, 2, 9, 18);
//		//row.addCell(cell);
//		//table.addRow(row);
		
		row = ExcelRow.createRow(sheet,3);
		ExcelRow.setRowHeight(row,1000);
		cell = ExcelCell.createDefaultCell(row,"项目", 0);
		ExcelCell.setMerge(cell, 3, 3, 0, 2);
//		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row,"人次", 3);
//		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row,"现金支出", 4);
//		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row,"账户支出", 5);
//		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row,"公务员补助支出", 6);
//		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row,"统筹支出", 7);
//		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row,"救助金支出", 8);
//		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row,"特殊人员统筹支出", 9);
//		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row,"二级保健对象补助支出", 10);
//		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row,"高知支出", 11);
//		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row,"一级保健对象补助支出", 12);
//		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row,"费用合计", 13);
//		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row,"偿付统筹\r\n(中心记账)", 14);//\r\n强制换行
//		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row,"住院均值或定额", 15);
//		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row,"审核扣款", 16);
//		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row,"预留风险基金费用", 17);
//		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row,"实际偿付金额", 18);
//		//row.addCell(cell);
//		//table.addRow(row);
		
		row = ExcelRow.createRow(sheet,4);
		cell = ExcelCell.createDefaultCell(row, "住院", 0);
		ExcelCell.setMerge(cell, 4, 6, 0, 0);
//		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row, "超起付", 1);
		ExcelCell.setMerge(cell, 4, 4, 1, 2);
//		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row, "258", 3);
//		//row.addCell(cell);
		for(int i = 4; i<=18; i++){
			cell = ExcelCell.createCell(row, "6,666.66", i, ExcelCellStyle.newInstance(handler,ExcelCellStyle.MONEY_STYLE));
//			//row.addCell(cell);
		}
//		//table.addRow(row);
		
		row = ExcelRow.createRow(sheet,5);
		cell = ExcelCell.createDefaultCell(row, "未超起付", 1);
		ExcelCell.setMerge(cell, 5, 5, 1, 2);
//		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row, "", 3);
//		//row.addCell(cell);
		for(int i = 4; i<=18; i++){
			cell = ExcelCell.createCell(row, "333.33",  i,ExcelCellStyle.newInstance(handler,ExcelCellStyle.MONEY_STYLE));
//			//row.addCell(cell);
		}
//		//table.addRow(row);
		
		sheet.addPageBreak(6);
		
		row = ExcelRow.createRow(sheet,6);
		cell = ExcelCell.createDefaultCell(row, "合计", 1);
		ExcelCell.setMerge(cell, 6, 6, 1, 2);
//		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row, "258", 3);
//		//row.addCell(cell);
		for(int i = 4; i<=18; i++){
			cell = ExcelCell.createCell(row, "999.99", i, ExcelCellStyle.newInstance(handler, ExcelCellStyle.MONEY_STYLE));
//			//row.addCell(cell);
		}
//		//table.addRow(row);
		
		row = ExcelRow.createRow(sheet, 7);
		cell = ExcelCell.createDefaultCell(row, "门诊", 0);
		ExcelCell.setMerge(cell,7, 7, 0, 2);
		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row, "7172", 3);
		//row.addCell(cell);
		for(int i = 4; i<=18; i++){
			cell = ExcelCell.createCell(row, "555.55", i, ExcelCellStyle.newInstance(handler,ExcelCellStyle.MONEY_STYLE));
			//row.addCell(cell);
		}
		//table.addRow(row);
		
		row = ExcelRow.createRow(sheet,8);
		cell = ExcelCell.createDefaultCell(row, "药店", 0);
		ExcelCell.setMerge(cell,8, 8, 0, 2);
		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row, "0", 3);
		//row.addCell(cell);
		for(int i = 4; i<=18; i++){
			cell = ExcelCell.createCell(row, "0.00", i, ExcelCellStyle.newInstance(handler,ExcelCellStyle.MONEY_STYLE));
			//row.addCell(cell);
		}
		//table.addRow(row);
		
		row = ExcelRow.createRow(sheet,9);
		cell = ExcelCell.createDefaultCell(row, "公务员", 0);
		ExcelCell.setMerge(cell,9, 12, 0, 0);
		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row, "工伤", 1);
		ExcelCell.setMerge(cell,9, 10, 1, 1);
		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row, "门诊", 2);
		//row.addCell(cell);
		for(int i = 3; i<=18; i++){
			cell = ExcelCell.createCell(row, "", i, ExcelCellStyle.newInstance(handler, ExcelCellStyle.MONEY_STYLE));
			//row.addCell(cell);
		}
		//table.addRow(row);
		
		row = ExcelRow.createRow(sheet, 10);
		cell = ExcelCell.createDefaultCell(row, "住院", 2);
		//row.addCell(cell);
		for(int i = 3; i<=18; i++){
			cell = ExcelCell.createCell(row, "", i, ExcelCellStyle.newInstance(handler,ExcelCellStyle.MONEY_STYLE));
			//row.addCell(cell);
		}
		//table.addRow(row);
		
		row = ExcelRow.createRow(sheet, 11);
		cell = ExcelCell.createDefaultCell(row, "生育", 1);
		ExcelCell.setMerge(cell,11, 12, 1, 1);
		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row, "门诊", 2);
		//row.addCell(cell);
		for(int i = 3; i<=18; i++){
			cell = ExcelCell.createCell(row, "", i, ExcelCellStyle.newInstance(handler, ExcelCellStyle.MONEY_STYLE));
			//row.addCell(cell);
		}
		//table.addRow(row);
		
		row = ExcelRow.createRow(sheet, 12);
		cell = ExcelCell.createDefaultCell(row, "住院", 2);
		//row.addCell(cell);
		for(int i = 3; i<=18; i++){
			cell = ExcelCell.createCell(row, "", i, ExcelCellStyle.newInstance(handler,ExcelCellStyle.MONEY_STYLE));
			//row.addCell(cell);
		}
		//table.addRow(row);
		
		row = ExcelRow.createRow(sheet,13);
		cell = ExcelCell.createDefaultCell(row, "合计", 0);
		ExcelCell.setMerge(cell,13, 13, 0, 2);
		//row.addCell(cell);
		cell = ExcelCell.createDefaultCell(row, "392", 3);
		//row.addCell(cell);
		for(int i = 4; i<=18; i++){
			cell = ExcelCell.createCell(row, "10000.00", i, ExcelCellStyle.newInstance(handler,ExcelCellStyle.MONEY_STYLE));
			//row.addCell(cell);
		}
		//table.addRow(row);
		
		row = ExcelRow.createRow(sheet,14);
		cell = ExcelCell.createCell(row, "实际拨付金额合计(大写人民币单位元)：壹佰柒拾肆万贰仟贰佰肆拾陆元伍角捌分", 0, ExcelCellStyle.newInstance(handler,ExcelCellStyle.DEFAULT_STYLE_LEFT));
		ExcelCell.setMerge(cell,14, 14, 0, 18);
		//row.addCell(cell);
		//table.addRow(row);
		
		handler.addSheet(sheet);
		try {
			handler.export(new FileOutputStream("D:\\test.xlsx"));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}
}
