package se.minstrel.tools.xssfbuilder.impl;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import se.minstrel.tools.xssfbuilder.SheetBuilder;
import se.minstrel.tools.xssfbuilder.WorkbookBuilder;
import se.minstrel.tools.xssfbuilder.markers.MarkerManager;

public class WorkbookBuilderImpl implements WorkbookBuilder {

	private XSSFWorkbook workbook;
	private Map <String, SheetBuilder> sheetBuilders = new HashMap <String, SheetBuilder> ();
	
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
			sheetBuilders.put(name, new SheetBuilderImpl(support, workbook.getSheetAt(i)));
		}
	}

	public SheetBuilder sheet(String name) {
		if (!sheetBuilders.containsKey(name)) {
			sheetBuilders.put(name, new SheetBuilderImpl(support, workbook.createSheet(name)));
		}
		return sheetBuilders.get(name);
	}
	
}
