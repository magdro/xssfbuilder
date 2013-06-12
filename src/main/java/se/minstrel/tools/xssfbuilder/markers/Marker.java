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

public class Marker {

    private String sheetName;
    private int row;
    private int col;
    private String name;

    public Marker(String name, String sheetName, int row, int col) {
	this.sheetName = sheetName;
	this.name = name;
	this.row = row;
	this.col = col;
    }

    public String getSheetName() {
	return sheetName;
    }

    public String getName() {
	return name;
    }

    public int getRow() {
	return row;
    }

    public int getCol() {
	return col;
    }

    @Override
    public int hashCode() {
	final int prime = 31;
	int result = 1;
	result = prime * result + col;
	result = prime * result + ((name == null) ? 0 : name.hashCode());
	result = prime * result + row;
	result = prime * result
		+ ((sheetName == null) ? 0 : sheetName.hashCode());
	return result;
    }

    @Override
    public boolean equals(Object obj) {
	if (this == obj)
	    return true;
	if (obj == null)
	    return false;
	if (getClass() != obj.getClass())
	    return false;
	Marker other = (Marker) obj;
	if (col != other.col)
	    return false;
	if (name == null) {
	    if (other.name != null)
		return false;
	} else if (!name.equals(other.name))
	    return false;
	if (row != other.row)
	    return false;
	if (sheetName == null) {
	    if (other.sheetName != null)
		return false;
	} else if (!sheetName.equals(other.sheetName))
	    return false;
	return true;
    }

}
