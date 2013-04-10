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

import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import se.minstrel.tools.xssfbuilder.ColumnBuilder;

public class ColumnBuilderImpl implements ColumnBuilder {

	private Support support;
	private XSSFSheet sheet;
	private int colNr;
	
	public ColumnBuilderImpl(Support support, XSSFSheet sheet, int colNr) {
		this.sheet = sheet;
		this.colNr = colNr;
		this.support = support;
	}

	@Override
	public ColumnBuilder autoWidth() {
		return autoWidth(false);
	}
	
	@Override
	public ColumnBuilder autoWidth(boolean evaluateFormulas) {
		if (evaluateFormulas) {
			support.getFormulaEvaluator().clearAllCachedResultValues();
			support.getFormulaEvaluator().evaluateAll();
		}
		
		sheet.autoSizeColumn(colNr);
		return this;
	}

}
