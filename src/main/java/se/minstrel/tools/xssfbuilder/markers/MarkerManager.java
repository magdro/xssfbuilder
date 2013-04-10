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
