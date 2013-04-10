package se.minstrel.tools.xssfbuilder;

import se.minstrel.tools.xssfbuilder.style.Style;

public interface CellBuilder {

	CellBuilder value(String value);

	CellBuilder value(Number value);

	CellBuilder formula(String formula);

	CellBuilder style(Style style);

	CellBuilder mark(String tag);

}
