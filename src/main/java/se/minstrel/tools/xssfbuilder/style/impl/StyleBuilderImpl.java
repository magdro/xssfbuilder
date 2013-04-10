package se.minstrel.tools.xssfbuilder.style.impl;

import java.awt.Color;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import se.minstrel.tools.xssfbuilder.impl.Support;
import se.minstrel.tools.xssfbuilder.style.Style;
import se.minstrel.tools.xssfbuilder.style.StyleBuilder;

public class StyleBuilderImpl implements StyleBuilder {

	private StyleImpl style = new StyleImpl();
	private Support support;
	
	public StyleBuilderImpl(Support support) {
		this.support = support;
	}
	
	public StyleBuilder italics() {
		style.setItalics(true);
		return this;
	}

	public StyleBuilder bgColor(Color color) {
		style.setBgColor(color);
		return this;
	}

	public StyleBuilder fgColor(Color color) {
		style.setFgColor(color);
		return this;
	}

	public StyleBuilder bold() {
		style.setBold(true);
		return this;
	}

	public Style apply() {
		if (!support.hasStyle(style)) {
			support.setStyle(style, buildStyle());
		}
		return style;
	}
	
	private XSSFCellStyle buildStyle() {
		XSSFWorkbook wb = support.getWorkbook();
		
		XSSFFont font = wb.createFont();
		font.setFontName(style.getFont());
		font.setFontHeight(style.getFontSize());
		font.setBold(style.isBold());
		font.setItalic(style.isItalics());
		font.setColor(new XSSFColor(style.getFgColor()));
		
		XSSFCellStyle cs = support.getWorkbook().createCellStyle();
		cs.setFont(font);
		
		cs.setFillForegroundColor(new XSSFColor(style.getBgColor()));
		cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		//style.getBgColor();
		//style.getFgColor();
		//style.isBold();
		//style.isItalics();
		
		return cs;
	}

	@Override
	public StyleBuilder font(String font) {
		style.setFont(font);
		return this;
	}

}
