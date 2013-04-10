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
		support.getMarkerManager().mark(cell.getSheet().getSheetName(), row, col, tag);
		return this;
	}

}
