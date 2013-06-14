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

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import se.minstrel.tools.xssfbuilder.SheetBuilder;
import se.minstrel.tools.xssfbuilder.WorkbookBuilder;
import se.minstrel.tools.xssfbuilder.markers.MarkerManager;
import se.minstrel.tools.xssfbuilder.style.StyleBuilder;
import se.minstrel.tools.xssfbuilder.style.impl.StyleBuilderImpl;

public class WorkbookBuilderImpl implements WorkbookBuilder {

    private XSSFWorkbook workbook;
    private Map<String, SheetBuilder> sheetBuilders = new HashMap<String, SheetBuilder>();

    private Support support;

    public WorkbookBuilderImpl(XSSFWorkbook workbook) {
	this.workbook = workbook;
	support = new Support(workbook, new MarkerManager());

	initFromWorkbook();
    }

    private void initFromWorkbook() {
	sheetBuilders.clear();
	for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
	    String name = workbook.getSheetName(i);
	    sheetBuilders.put(name,
		    new SheetBuilderImpl(support, workbook.getSheetAt(i)));
	}
    }

    public SheetBuilder sheet(String name) {
	if (!sheetBuilders.containsKey(name)) {
	    sheetBuilders.put(name,
		    new SheetBuilderImpl(support, workbook.createSheet(name)));
	}
	return sheetBuilders.get(name);
    }

    @Override
    public StyleBuilder styleBuilder() {
	return new StyleBuilderImpl(support);
    }

}
