package br.com.dts.excelence.range;

import br.com.dts.excelence.ExcelCell;
import br.com.dts.excelence.ExcelSheet;
import br.com.dts.excelence.style.ExcelStyle;

public class ExcelVerticalRange extends ExcelLinearRange {

	//
	//  Constructors
	//
	
	public ExcelVerticalRange(ExcelSheet sheet, int col, int beginRow, int endRow) {
		super(sheet, col, beginRow, endRow);

	}
	
	//
	//  Getters and Setters
	//
	
	public int col() {
		return i();
	}
	
	public int beginRow() {
		return beginJ();
	}
	
	public int endRow() {
		return endJ();
	}
	
	//
	//  Public Methods
	//

	public ExcelVerticalRange style(ExcelStyle... styles) {

		return this;
	}

	@Override
	public ExcelCell cell(int row) {
		return new ExcelCell(sheet(), row, col());
	}

}
