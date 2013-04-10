package se.minstrel.tools.xssfbuilder;

import se.minstrel.tools.xssfbuilder.style.Style;
import se.minstrel.tools.xssfbuilder.style.StyleBuilder;

public interface SheetBuilder {

	RowBuilder row(int rowNr);

	ColumnBuilder col(int colNr);

	CellBuilder cell(int rowNr, int colNr);

	FormulaBuilder formula();

	StyleBuilder styleBuilder();

	SheetBuilder applyAreaStyle(int rowStart, int rowEnd, int colStart, int colEnd, Style style);

	SheetBuilder clearMarkers();

}
