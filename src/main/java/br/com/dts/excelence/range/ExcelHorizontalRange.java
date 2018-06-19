package br.com.dts.excelence.range;

import br.com.dts.excelence.ExcelSheet;
import br.com.dts.excelence.style.ExcelStyle;

public class ExcelHorizontalRange extends ExcelLinearRange {

	//
	// Constructors
	//
	
	public ExcelHorizontalRange(ExcelSheet sheet, int row, int beginColumn, int endColumn) {
		super(sheet, row, beginColumn, endColumn);

	}
	
	//
	//  Getters and Setters
	//
	
	public int row() {
		return i();
	}
	
	public int beginCol() {
		return beginJ();
	}
	
	public int endCol() {
		return endJ();
	}
	
	//
	//  Public Methods
	//

	public ExcelHorizontalRange style(ExcelStyle... styles) {

		return this;
	}

	@Override
	public ExcelCell cell(int col) {
		return new ExcelCell(sheet(), row(), col);
	}

}
