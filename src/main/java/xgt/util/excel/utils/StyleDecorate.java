/**
 * 
 */
package xgt.util.excel.utils;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;

import xgt.util.excel.styles.Border;
import xgt.util.excel.styles.Directive;

/**
 * @author Gavin
 *
 */
public class StyleDecorate {
	
	private StyleDecorate(){};
	
	public static CellStyle addBorder(CellStyle style){
		Border border = Border.newDefaultInstance();
		style = addBorder(style, border);
		return style;
	}
	
	public static CellStyle addBorder(CellStyle style,Border border){
		
		style.setBorderBottom(border.getType(Directive.BOTTOM));
		style.setBorderLeft(border.getType(Directive.LEFT));
		style.setBorderRight(border.getType(Directive.RIGHT));
		style.setBorderTop(border.getType(Directive.TOP));
		
		style.setBottomBorderColor(border.getColor(Directive.BOTTOM));
		style.setLeftBorderColor(border.getColor(Directive.LEFT));
		style.setRightBorderColor(border.getColor(Directive.RIGHT));
		style.setTopBorderColor(border.getColor(Directive.TOP));
		
		return style;
	}
	
	public static CellStyle decorateLeft(CellStyle style){
		return decorateAlign(style,CellStyle.ALIGN_LEFT);
	}
	
	public static CellStyle decorateCenter(CellStyle style){
		return decorateAlign(style,CellStyle.ALIGN_CENTER);
	}

	public static CellStyle decorateRight(CellStyle style){
		return decorateAlign(style,CellStyle.ALIGN_RIGHT);
	}
	
	public static CellStyle decorateAlign(CellStyle style,short align){
		style.setAlignment(align);
		return style;
	}
	
	public static CellStyle decorateJustify(CellStyle style){
		style.setAlignment(CellStyle.ALIGN_JUSTIFY);
		return style;
	}
	
	public static CellStyle decorateAsTitle(CellStyle style, Font font){
		style.setFont(FontDecorate.decorateDefaultTitle(font));
		style.setAlignment(CellStyle.ALIGN_CENTER);
		return style;
	}
	
	public static CellStyle decorateAsHeader(CellStyle style, Font font){
		style.setFont(FontDecorate.decorateBold(font));
		style.setAlignment(CellStyle.ALIGN_CENTER);
		return style;
	}
	
	public static CellStyle decorateAsFooter(CellStyle style, Font font){
		style.setFont(FontDecorate.decorateBold(font));
		style.setAlignment(CellStyle.ALIGN_CENTER);
		return style;
	}
	
	public static CellStyle decorateAs$(CellStyle style,DataFormat format){
		style.setAlignment(CellStyle.ALIGN_RIGHT);
        style.setDataFormat(format.getFormat("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)"));
		return style;
	}
	
	public static CellStyle decorateAs￥(CellStyle style,DataFormat format){
		style.setAlignment(CellStyle.ALIGN_RIGHT);
        style.setDataFormat(format.getFormat("_(￥* #,##0.00_);_(￥* (#,##0.00);_(￥* \"-\"??_);_(@_)"));
		return style;
	}
	
	public static CellStyle decorateAsPercentate(CellStyle style,DataFormat format){
		style.setAlignment(CellStyle.ALIGN_RIGHT);
        style.setDataFormat(format.getFormat("0.00%"));
		return style;
	}
	
	public static CellStyle decorateAsDate(CellStyle style,DataFormat format){
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setDataFormat(format.getFormat("yyyy-MM-dd"));
		return style;
	}
	
	public static CellStyle decorateBgColor(CellStyle style, short color){
		style.setFillForegroundColor(color);
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		return style;
	}
	
	public static CellStyle decorateBgYellow(CellStyle style){
		return decorateBgColor(style,IndexedColors.YELLOW.getIndex());
	}
	
	public static CellStyle decorateBgGrey(CellStyle style){
		return decorateBgColor(style,IndexedColors.GREY_25_PERCENT.getIndex());
	}
	
	public static CellStyle decorateBgBlue(CellStyle style){
		return decorateBgColor(style,IndexedColors.BLUE.getIndex());
	}
	
}
