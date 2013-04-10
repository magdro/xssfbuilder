package se.minstrel.tools.xssfbuilder.markers;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;


public class MarkerManager {

	private Map<String, List<Marker>> markers = new HashMap<String, List<Marker>>();

	public void mark(String sheet, int row, int col, String tag) {
		Marker marker = new Marker(tag, sheet, row, col);

		if (!markers.containsKey(tag)) {
			markers.put(tag, new ArrayList<Marker>());
		}

		List<Marker> list = markers.get(tag);
		if (!list.contains(marker)) {
			list.add(marker);
		}
	}

	public void unmark(String sheet, int row, int col, String tag) {
		if (markers.containsKey(tag)) {
			Marker marker = new Marker(tag, sheet, row, col);
			List <Marker> list = markers.get(tag);
			list.remove(marker);
		}
	}

	public void unmark(String sheet, int row, int col) {
		Iterator <Entry <String, List <Marker>>> outer = markers.entrySet().iterator();
		while (outer.hasNext()) {
			Iterator <Marker> inner = outer.next().getValue().iterator();
			while (inner.hasNext()) {
				Marker marker = inner.next();
				if (marker.getCol() == col && marker.getRow() == row && marker.getSheetName().compareTo(sheet) == 0) {
					inner.remove();
				}
			}
		}
	}

	public void unmark(String tag) {
		markers.remove(tag);
	}

	public void unmark() {
		markers.clear();
	}
	
	public List <Marker> get(String tag) {
		return markers.containsKey(tag) ? markers.get(tag) : new ArrayList <Marker> ();
	}
}
