package xgt.util.easyexcel.last;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelCellStyle {

	/**
	 * 字体加粗
	 */
	public static final short BOLD_ADD = Font.BOLDWEIGHT_BOLD;
	/**
	 * 字体不加粗
	 */
	public static final short BOLD_NOMAL = Font.BOLDWEIGHT_NORMAL;
	/**
	 * 红色字体
	 */
	public static final short COLOR_RED = Font.COLOR_RED;
	/**
	 * 黑色字体
	 */
	public static final short COLOR_NORMAL = Font.COLOR_NORMAL;
	/**
	 * 字体加下划线
	 */
	public static final byte UNDERLINE_ONE = Font.U_SINGLE_ACCOUNTING;
	/**
	 * 字体不加下划线
	 */
	public static final byte UNDERLINE_NONE = Font.U_NONE;
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

	public static final short FONT_SIZE_BIGGER = 16;

	public static final short FONT_SIZE_BIG = 14;

	public static final short FONT_SIZE_NORMAL = 12;

	public static final short FONT_SIZE_SMALL = 10;
	/**
	 * head样式
	 */
	public static final int HEAD_STYLE = 1;
	/**
	 * money样式
	 */
	public static final int MONEY_STYLE = 2;
	/**
	 * 不带边框的居中
	 */
	public static final int BORDER_NONE_CENTER = 3;
	/**
	 * 不带边框的居左
	 */
	public static final int BORDER_NONE_LEFT = 4;
	/**
	 * 不带边框的居右
	 */
	public static final int BORDER_NONE_RIGHT = 5;
	/**
	 * 普通样式居左，带边框
	 */
	public static final int DEFAULT_STYLE_LEFT = 6;
	/**
	 * 普通样式居中，带边框
	 */
	public static final int DEFAULT_STYLE_CENTER = 0;
	
	//导出一次内存中的样式
	private static Map<String,CellStyle> styles = new HashMap<String,CellStyle>();
	
	private ExcelCellStyle(){}
	/**
	 * 创建单元格样式
	 * @param fontColor 字体颜色
	 * @param fontBlod 字体加粗值
	 * @param fontUnderLine 字体下划线
	 * @param fontSize 字体大小
	 * @param fontAlign 字体对齐方式
	 * @param isBorder 单元格是否带边框
	 * @return
	 */
	public static CellStyle newInstance(ExcelHandler handler, short fontColor, short fontBlod,
			byte fontUnderLine, short fontSize, short fontAlign,
			boolean isBorder){
		return newInstance(handler.getWorkbook(), fontColor, fontBlod, fontUnderLine, fontSize, fontAlign, isBorder);
	}
	/**
	 * 
	 * @param wb
	 * @param fontColor
	 * @param fontBlod
	 * @param fontUnderLine
	 * @param fontSize
	 * @param fontAlign
	 * @param isBorder
	 * @return
	 */
	private static CellStyle newInstance(Workbook wb, short fontColor, short fontBlod,
			byte fontUnderLine, short fontSize, short fontAlign,
			boolean isBorder){
		StringBuffer key = new StringBuffer();
		key.append(fontColor).append(fontBlod).append(fontUnderLine).append(fontSize).append(fontAlign).append(isBorder);
		CellStyle style = styles.get(key.toString());
		if(style==null){
			style = ExcelUtil.createCellStyle(wb, fontColor, fontBlod, fontUnderLine, fontSize, fontAlign, isBorder);
			styles.put(key.toString(), style);
		}
		return style;
	}
	
	/**
	 * @see newStyle
	 * 默认单元格样式，带边框普通字体居中
	 * @return
	 */
	public static CellStyle newDefaultInstance(Workbook wb) {
			return newInstance(wb,ExcelUtil.FONT_COLOR_NORMAL,ExcelUtil.FONT_BOLD_NOMAL,
					ExcelUtil.NONE_UNDERLINE,ExcelUtil.FONT_NORMAL,ExcelUtil.ALIGN_CENTER,true);
	}
	
	/**
	 * @see newStyle
	 * 默认单元格样式，带边框普通字体居中
	 * @return
	 */
	public static CellStyle newDefaultInstance(ExcelHandler handler) {
			return newInstance(handler,ExcelUtil.FONT_COLOR_NORMAL,ExcelUtil.FONT_BOLD_NOMAL,
					ExcelUtil.NONE_UNDERLINE,ExcelUtil.FONT_NORMAL,ExcelUtil.ALIGN_CENTER,true);
	}
	
	/**
	 * 带边框
	 * @param handler
	 * @param align 字体对齐方式
	 * @return
	 */
	public static CellStyle newInstance(ExcelHandler handler,short align) {
			return newInstance(handler,ExcelUtil.FONT_COLOR_NORMAL,ExcelUtil.FONT_BOLD_NOMAL,
					ExcelUtil.NONE_UNDERLINE,ExcelUtil.FONT_NORMAL,align,true);
	}
	
	/**
	 * @see newStyle
	 * 创建单元格样式
	 * @param style 样式类型
	 * @return
	 */
	public static CellStyle newInstance(ExcelHandler handler,int style){
		switch(style){
			case 0 : 
				return newDefaultInstance(handler);
			case 1 : 
				return newInstance(
						handler,
						ExcelUtil.FONT_COLOR_NORMAL,
						ExcelUtil.FONT_BOLD,
						ExcelUtil.NONE_UNDERLINE,
						ExcelUtil.FONT_NORMAL,
						ExcelUtil.ALIGN_CENTER,
						true);
			case 2 : 
				return newInstance(
						handler,
						ExcelUtil.FONT_COLOR_NORMAL,
						ExcelUtil.FONT_BOLD_NOMAL,
						ExcelUtil.NONE_UNDERLINE,
						ExcelUtil.FONT_NORMAL,
						ExcelUtil.ALIGN_RIGHT,
						true
					);
			case 3 : 
				return newInstance(
						handler,
						ExcelUtil.FONT_COLOR_NORMAL,
						ExcelUtil.FONT_BOLD_NOMAL,
						ExcelUtil.NONE_UNDERLINE,
						ExcelUtil.FONT_NORMAL,
						ExcelUtil.ALIGN_CENTER,
						false
					);
			case 4 : 
				return newInstance(
						handler,
						ExcelUtil.FONT_COLOR_NORMAL,
						ExcelUtil.FONT_BOLD_NOMAL,
						ExcelUtil.NONE_UNDERLINE,
						ExcelUtil.FONT_NORMAL,
						ExcelUtil.ALIGN_LEFT,
						false
					);
			case 5 : 
				return newInstance(
						handler,
						ExcelUtil.FONT_COLOR_NORMAL,
						ExcelUtil.FONT_BOLD_NOMAL,
						ExcelUtil.NONE_UNDERLINE,
						ExcelUtil.FONT_NORMAL,
						ExcelUtil.ALIGN_RIGHT,
						false
					);
			case 6 : 
				return newInstance(
						handler,
						ExcelUtil.FONT_COLOR_NORMAL,
						ExcelUtil.FONT_BOLD_NOMAL,
						ExcelUtil.NONE_UNDERLINE,
						ExcelUtil.FONT_NORMAL,
						ExcelUtil.ALIGN_LEFT,
						true
					);
			default : 
				return newDefaultInstance(handler);
		}
	}
	
	public static void clear(){
		styles.clear();
	}
}
