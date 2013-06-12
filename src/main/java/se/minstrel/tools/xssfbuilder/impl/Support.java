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

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import se.minstrel.tools.xssfbuilder.markers.MarkerManager;
import se.minstrel.tools.xssfbuilder.style.Style;

public class Support {

    private MarkerManager markerManager;
    private Map<Style, XSSFCellStyle> styleMap;
    private XSSFWorkbook workbook;
    private XSSFDataFormat dataFormat;
    private FormulaEvaluator formulaEvaluator;

    public Support(XSSFWorkbook workbook, MarkerManager markerManager) {
	this.markerManager = markerManager;
	this.workbook = workbook;
	this.styleMap = new HashMap<Style, XSSFCellStyle>();
	this.dataFormat = workbook.createDataFormat();
	this.formulaEvaluator = workbook.getCreationHelper()
		.createFormulaEvaluator();
    }

    public FormulaEvaluator getFormulaEvaluator() {
	return formulaEvaluator;
    }

    public XSSFDataFormat getDataFormat() {
	return dataFormat;
    }

    public MarkerManager getMarkerManager() {
	return markerManager;
    }

    public XSSFCellStyle getStyle(Style style) {
	return styleMap.get(style);
    }

    public void setStyle(Style style, XSSFCellStyle cellStyle) {
	styleMap.put(style, cellStyle);
    }

    public boolean hasStyle(Style style) {
	return styleMap.containsKey(style);
    }

    public XSSFWorkbook getWorkbook() {
	return workbook;
    }
}
