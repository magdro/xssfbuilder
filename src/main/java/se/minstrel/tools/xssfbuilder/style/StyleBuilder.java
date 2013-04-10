package se.minstrel.tools.xssfbuilder.style;

import java.awt.Color;

public interface StyleBuilder {

	StyleBuilder italics();

	StyleBuilder bgColor(Color color);
	
	StyleBuilder fgColor(Color color);

	StyleBuilder bold();
	
	StyleBuilder font(String font);

	Style apply();

}
