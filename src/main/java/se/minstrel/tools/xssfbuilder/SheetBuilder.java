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

import se.minstrel.tools.xssfbuilder.style.Style;
import se.minstrel.tools.xssfbuilder.style.StyleBuilder;

public interface SheetBuilder {

    RowBuilder row(int rowNr);

    ColumnBuilder col(int colNr);

    CellBuilder cell(int rowNr, int colNr);

    FormulaBuilder formula();

    StyleBuilder styleBuilder();

    SheetBuilder applyAreaStyle(int rowStart, int rowEnd, int colStart,
	    int colEnd, Style style);

    SheetBuilder clearMarkers();

    SheetBuilder autoWidth();

    SheetBuilder autoWidth(boolean evaluateFormulas);

}
