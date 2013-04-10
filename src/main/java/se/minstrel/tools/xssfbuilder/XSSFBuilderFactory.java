package se.minstrel.tools.xssfbuilder;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import se.minstrel.tools.xssfbuilder.impl.WorkbookBuilderImpl;

public class XSSFBuilderFactory {

	public static WorkbookBuilder createBuilder(XSSFWorkbook workbook) {
		return new WorkbookBuilderImpl(workbook);
	}

}
