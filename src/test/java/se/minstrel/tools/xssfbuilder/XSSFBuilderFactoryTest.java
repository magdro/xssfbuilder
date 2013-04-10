package se.minstrel.tools.xssfbuilder;

import java.awt.Color;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class XSSFBuilderFactoryTest {

	@Test
	public void test() {
		XSSFWorkbook xwb = new XSSFWorkbook();
		
		WorkbookBuilder wb = XSSFBuilderFactory.createBuilder(xwb);
		SheetBuilder sheet = wb.sheet("sheet1");
		
		sheet.clearMarkers();
		
		sheet.cell(0, 0).value("Column 1").style(sheet.styleBuilder().bold().bgColor(Color.GRAY).apply());
		sheet.cell(0, 1).value("Column 2").style(sheet.styleBuilder().bold().bgColor(Color.GRAY).apply());
		sheet.cell(0, 2).value("Column 3").style(sheet.styleBuilder().bold().bgColor(Color.GRAY).apply());
		
		for (int r = 1; r <= 3; r++) {
			for (int c = 0; c < 3; c++) {
				sheet.cell(r, c).value(new Double(r*c));
			}
		}
		
		sheet.cell(4, 0).formula(sheet.formula().sum(1, 3, 0, 0)).mark("Mark1");
		sheet.cell(4, 1).formula(sheet.formula().sum(1, 3, 1, 1)).mark("Mark1");
		sheet.cell(4, 2).formula(sheet.formula().sum(1, 3, 2, 2)).mark("Mark1");
		
		sheet.cell(4, 3).formula(sheet.formula().sumMarkers("Mark1"));
		
		sheet.applyAreaStyle(4, 4, 0, 3, sheet.styleBuilder().italics().bgColor(Color.BLACK).fgColor(Color.YELLOW).apply());

		try {
			write(xwb, "out.xlsx");
		} catch (IOException ioe) {
			ioe.printStackTrace();
		}
	}
	
	private void write(XSSFWorkbook workbook, String filename) throws IOException {
		try (FileOutputStream os = new FileOutputStream(filename)) {
			workbook.write(os);
			os.flush();
		}
	}
}
