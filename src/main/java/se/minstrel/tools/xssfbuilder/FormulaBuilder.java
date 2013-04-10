package se.minstrel.tools.xssfbuilder;

public interface FormulaBuilder {

	String sum(int rowStart, int rowEnd, int colStart, int colEnd);

	String sumMarkers(String marker);


}
