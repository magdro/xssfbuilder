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
package se.minstrel.tools.xssfbuilder;

import java.awt.Color;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class XSSFBuilderFactoryTest {

    @Test
    public void test() {
	XSSFWorkbook xwb = new XSSFWorkbook();

	WorkbookBuilder wb = XSSFBuilderFactory.createBuilder(xwb);
	wb.styleBuilder().fontSize((short)8).makeDefault();

	SheetBuilder sheet = wb.sheet("sheet1");

	sheet.clearMarkers();

	sheet.cell(0, 0).value("Column 1")
		.style(sheet.styleBuilder().bold().bgColor(Color.GRAY).apply());
	sheet.cell(0, 1).value("Column 2")
		.style(sheet.styleBuilder().bold().bgColor(Color.GRAY).apply());
	sheet.cell(0, 2).value("Column 3")
		.style(sheet.styleBuilder().bold().bgColor(Color.GRAY).apply());

	for (int r = 1; r <= 3; r++) {
	    for (int c = 0; c < 3; c++) {
		sheet.cell(r, c).value(new Double(r * c));
	    }
	}

	sheet.cell(4, 0).formula(sheet.formula().sum(1, 3, 0, 0)).mark("Mark1");
	sheet.cell(4, 1).formula(sheet.formula().sum(1, 3, 1, 1)).mark("Mark1");
	sheet.cell(4, 2).formula(sheet.formula().sum(1, 3, 2, 2)).mark("Mark1");

	sheet.cell(4, 3).formula(sheet.formula().sumMarkers("Mark1"));

	sheet.applyAreaStyle(4, 4, 0, 3, sheet.styleBuilder().italics()
		.bgColor(Color.BLACK).fgColor(Color.YELLOW).apply());

	sheet.cell(6, 0).value("Todays date")
		.style(sheet.styleBuilder().bold().apply());
	sheet.cell(6, 1).value(new Date())
		.style(sheet.styleBuilder().format("yyyy-mm-dd").apply());

	sheet.col(0).autoWidth();
	sheet.autoWidth(true);

	try {
	    write(xwb, "out.xlsx");
	} catch (IOException ioe) {
	    ioe.printStackTrace();
	}
    }

    private void write(XSSFWorkbook workbook, String filename)
	    throws IOException {
	try (FileOutputStream os = new FileOutputStream(filename)) {
	    workbook.write(os);
	    os.flush();
	}
    }
}
