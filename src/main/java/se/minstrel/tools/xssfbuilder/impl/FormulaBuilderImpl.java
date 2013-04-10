package se.minstrel.tools.xssfbuilder.impl;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

import se.minstrel.tools.xssfbuilder.FormulaBuilder;
import se.minstrel.tools.xssfbuilder.markers.Marker;

public class FormulaBuilderImpl implements FormulaBuilder {

	private Support support;
	
	public FormulaBuilderImpl(Support support) {
		this.support = support;
	}
	
	public String sum(int rowStart, int rowEnd, int colStart, int colEnd) {
		CellRangeAddress cra = new CellRangeAddress(rowStart, rowEnd, colStart, colEnd);
		return String.format("SUM(%s)", cra.formatAsString());
	}

	public String sumMarkers(String marker) {
		List <String> refs = new ArrayList <String> ();
		for (Marker m : support.getMarkerManager().get(marker)) {
			refs.add(String.format("%s!%s", m.getSheetName(), new CellReference(m.getRow(), m.getCol()).formatAsString()));
		}
		StringBuilder sb = new StringBuilder("SUM(");
		Iterator <String> it = refs.iterator();
		while (it.hasNext()) {
			sb.append(it.next());
			if (it.hasNext()) {
				sb.append(",");
			}
		}
		sb.append(")");
		return sb.toString();
	}

}
