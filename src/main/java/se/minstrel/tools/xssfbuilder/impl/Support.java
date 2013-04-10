package se.minstrel.tools.xssfbuilder.impl;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import se.minstrel.tools.xssfbuilder.markers.MarkerManager;
import se.minstrel.tools.xssfbuilder.style.Style;

public class Support {

	private MarkerManager markerManager;
	private Map<Style, XSSFCellStyle> styleMap;
	private XSSFWorkbook workbook;

	public Support(XSSFWorkbook workbook, MarkerManager markerManager) {
		this.markerManager = markerManager;
		this.workbook = workbook;
		this.styleMap = new HashMap <Style, XSSFCellStyle> ();
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
