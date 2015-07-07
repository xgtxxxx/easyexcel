/**
 * 
 */
package xgt.util.excel.utils;

import org.apache.poi.ss.usermodel.Font;

/**
 * @author Gavin
 *
 */
public class FontDecorate {
	
	public static final short FONT_SIZE_BIGGER = 16;

	public static final short FONT_SIZE_BIG = 14;

	public static final short FONT_SIZE_NORMAL = 12;

	public static final short FONT_SIZE_SMALL = 10;
	
	public static Font decorateColor(Font font,short color){
		font.setColor(color);
		return font;
	}
	
	public static Font decorateRed(Font font){
		return decorateColor(font, Font.COLOR_RED);
	}
	
	public static Font decorateUnderline(Font font, byte underline){
		font.setUnderline(underline);
		return font;
	}
	
	public static Font decorateDefaultUnderline(Font font){
		return decorateUnderline(font,Font.U_SINGLE);
	}
	
	public static Font decorateDefaultFamily(Font font){
		font.setFontName("宋体");
		return font;
	}
	
	public static Font decorateWeight(Font font,short weight){
		font.setBoldweight(weight);
		return font;
	}
	
	public static Font decorateSize(Font font,short size){
		font.setFontHeightInPoints(size);
		return font;
	}
	
	public static Font decorateBigSize(Font font){
		return decorateSize(font, FONT_SIZE_BIG);
	}
	
	public static Font decorateBold(Font font){
		return decorateWeight(font, Font.BOLDWEIGHT_BOLD);
	}
	
	public static Font decorateRedBold(Font font){
		font = decorateBold(font);
		font = decorateRed(font);
		return font;
	}
	
	public static Font decorateRedBoldUnderline(Font font){
		font = decorateBold(font);
		font = decorateRed(font);
		font = decorateDefaultUnderline(font);
		return font;
	}
	
	public static Font decorateRedUnderline(Font font){
		font = decorateRed(font);
		font = decorateDefaultUnderline(font);
		return font;
	}
	
	public static Font decorateBoldUnderline(Font font){
		font = decorateBold(font);
		font = decorateDefaultUnderline(font);
		return font;
	}
	
	public static Font decorateDefaultTitle(Font font){
		font = decorateBold(font);
		font = decorateBigSize(font);
		return font;
	}
}	
