/*
 * XSSFBuilder - an API making work with poi spreadsheets more time efficient.
 * (C) 2013, Magnus Drougge <magnus.drougge@gmail.com>
 *
 * This file is part of XSSFBuilder.
 *
 * XSSFBuilder is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * XSSFBuilder is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with XSSFBuilder.  If not, see <http://www.gnu.org/licenses/>.
 */
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

    public StyleBuilder makeDefault() {
	Style style = apply();
	support.setDefaultStyle(style);
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
	font.setFontHeightInPoints(style.getFontSize());
	font.setBold(style.isBold());
	font.setItalic(style.isItalics());
	if (style.getFgColor() != null) {
	    font.setColor(new XSSFColor(style.getFgColor()));
	}

	XSSFCellStyle cs = support.getWorkbook().createCellStyle();
	cs.setFont(font);

	if (style.getBgColor() != null) {
	    cs.setFillForegroundColor(new XSSFColor(style.getBgColor()));
	    cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	}

	if (style.getFormat() != null) {
	    cs.setDataFormat(support.getDataFormat().getFormat(
		    style.getFormat()));
	}
	// style.getBgColor();
	// style.getFgColor();
	// style.isBold();
	// style.isItalics();

	return cs;
    }

    @Override
    public StyleBuilder font(String font) {
	style.setFont(font);
	return this;
    }

    @Override
    public StyleBuilder format(String format) {
	style.setFormat(format);
	return this;
    }

    @Override
    public StyleBuilder base(Style style) {
	StyleImpl si = (StyleImpl) style;
	this.style.setBgColor(si.getBgColor());
	this.style.setBold(si.isBold());
	this.style.setFgColor(si.getFgColor());
	this.style.setFont(si.getFont());
	this.style.setFontSize(si.getFontSize());
	this.style.setFormat(si.getFormat());
	this.style.setItalics(si.isItalics());
	return this;
    }
}
