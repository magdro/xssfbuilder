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
