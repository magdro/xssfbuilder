package se.minstrel.tools.xssfbuilder.impl;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import se.minstrel.tools.xssfbuilder.CellBuilder;
import se.minstrel.tools.xssfbuilder.ColumnBuilder;
import se.minstrel.tools.xssfbuilder.FormulaBuilder;
import se.minstrel.tools.xssfbuilder.RowBuilder;
import se.minstrel.tools.xssfbuilder.SheetBuilder;
import se.minstrel.tools.xssfbuilder.style.Style;
import se.minstrel.tools.xssfbuilder.style.StyleBuilder;
import se.minstrel.tools.xssfbuilder.style.impl.StyleBuilderImpl;

public class SheetBuilderImpl implements SheetBuilder {

	private Support support;
	private XSSFSheet sheet;
	
	public SheetBuilderImpl(Support support, XSSFSheet sheet) {
		this.sheet = sheet;
		this.support = support;
	}
	
	public RowBuilder row(int rowNr) {
		XSSFRow row = sheet.getRow(rowNr);
		if (row == null) {
			row = sheet.createRow(rowNr);
		}
		
		return new RowBuilderImpl(rowNr);
	}

	public ColumnBuilder col(int colNr) {
		return new ColumnBuilderImpl(colNr);
	}

	public CellBuilder cell(int rowNr, int colNr) {
		XSSFRow row = sheet.getRow(rowNr);
		if (row == null) {
			row = sheet.createRow(rowNr);
		}
		
		XSSFCell cell = row.getCell(colNr);
		if (cell == null) {
			cell = row.createCell(colNr);
		}
		
		return new CellBuilderImpl(support, cell, rowNr, colNr);
	}

	public FormulaBuilder formula() {
		return new FormulaBuilderImpl(support);
	}

	public StyleBuilder styleBuilder() {
		return new StyleBuilderImpl(support);
	}

	public SheetBuilder applyAreaStyle(int rowStart, int rowEnd, int colStart,
			int colEnd, Style style) {
		for (int r = rowStart; r <= rowEnd; r++) {
			for (int c = colStart; c <= colEnd; c++) {
				cell(r, c).style(style);
			}
		}
		return this;
	}

	public SheetBuilder clearMarkers() {
		support.getMarkerManager().unmark();
		return this;
	}

}
