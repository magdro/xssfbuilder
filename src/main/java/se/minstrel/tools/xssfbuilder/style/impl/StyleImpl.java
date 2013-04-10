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
package se.minstrel.tools.xssfbuilder.style.impl;

import java.awt.Color;

import se.minstrel.tools.xssfbuilder.style.Style;

public class StyleImpl implements Style {

	private boolean bold = false;
	private boolean italics = false;
	private Color fgColor = Color.black;
	private Color bgColor = Color.white;
	private String font = "Calibri";
	private double fontSize = 11.0;

	public boolean isBold() {
		return bold;
	}

	public void setBold(boolean bold) {
		this.bold = bold;
	}

	public boolean isItalics() {
		return italics;
	}

	public void setItalics(boolean italics) {
		this.italics = italics;
	}

	public Color getFgColor() {
		return fgColor;
	}

	public void setFgColor(Color fgColor) {
		this.fgColor = fgColor;
	}

	public Color getBgColor() {
		return bgColor;
	}

	public void setBgColor(Color bgColor) {
		this.bgColor = bgColor;
	}

	public String getFont() {
		return font;
	}

	public void setFont(String font) {
		this.font = font;
	}

	public double getFontSize() {
		return fontSize;
	}

	public void setFontSize(double fontSize) {
		this.fontSize = fontSize;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((bgColor == null) ? 0 : bgColor.hashCode());
		result = prime * result + (bold ? 1231 : 1237);
		result = prime * result + ((fgColor == null) ? 0 : fgColor.hashCode());
		result = prime * result + ((font == null) ? 0 : font.hashCode());
		long temp;
		temp = Double.doubleToLongBits(fontSize);
		result = prime * result + (int) (temp ^ (temp >>> 32));
		result = prime * result + (italics ? 1231 : 1237);
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
		StyleImpl other = (StyleImpl) obj;
		if (bgColor == null) {
			if (other.bgColor != null)
				return false;
		} else if (!bgColor.equals(other.bgColor))
			return false;
		if (bold != other.bold)
			return false;
		if (fgColor == null) {
			if (other.fgColor != null)
				return false;
		} else if (!fgColor.equals(other.fgColor))
			return false;
		if (font == null) {
			if (other.font != null)
				return false;
		} else if (!font.equals(other.font))
			return false;
		if (Double.doubleToLongBits(fontSize) != Double
				.doubleToLongBits(other.fontSize))
			return false;
		if (italics != other.italics)
			return false;
		return true;
	}

}
