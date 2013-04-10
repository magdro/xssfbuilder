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
package se.minstrel.tools.xssfbuilder.impl;

import org.apache.poi.xssf.usermodel.XSSFCell;

import se.minstrel.tools.xssfbuilder.CellBuilder;
import se.minstrel.tools.xssfbuilder.style.Style;

public class CellBuilderImpl implements CellBuilder {

	private XSSFCell cell;
	private int row;
	private int col;

	private Support support;

	public CellBuilderImpl(Support support, XSSFCell cell, int row, int col) {
		this.cell = cell;
		this.row = row;
		this.col = col;
		this.support = support;
	}

	public CellBuilder value(String value) {
		cell.setCellValue(value);
		return this;
	}

	public CellBuilder value(Number value) {
		cell.setCellValue(value.doubleValue());
		return this;
	}

	public CellBuilder formula(String formula) {
		cell.setCellFormula(formula);
		return this;
	}

	public CellBuilder style(Style style) {
		cell.setCellStyle(support.getStyle(style));
		return this;
	}

	public CellBuilder mark(String tag) {
		support.getMarkerManager().mark(cell.getSheet().getSheetName(), row,
				col, tag);
		return this;
	}

}
