package br.com.dts.excelence;

import org.apache.poi.ss.usermodel.Sheet;

import br.com.dts.excelence.range.ExcelCell;
import br.com.dts.excelence.range.ExcelHorizontalRange;
import br.com.dts.excelence.range.ExcelSquareRange;
import br.com.dts.excelence.range.ExcelVerticalRange;

public class ExcelSheet {
	private final ExcelWorkbook workbook;
	private final Sheet poiSheet;
	
	//
	//  Constructors
	//
	
	public ExcelSheet(ExcelWorkbook workbook, String sheetName) {
		this.workbook = workbook;
		poiSheet = workbook.poiWorkbook().createSheet(sheetName);
	}
	
	//
	//  Getters and Setters
	//
	
	public ExcelWorkbook workbook() {
		return workbook;
	}
	
	public String sheetName() {
		return poiSheet.getSheetName();
	}
	
	public Sheet poiSheet() {
		return poiSheet;
	}
	
	//
	//  Public Methods
	//
	
	public ExcelCell cell(int row, int col) {
		return new ExcelCell(this, row, col);
	}

	public ExcelSquareRange range(int beginRow, int endRow, int beginColumn, int endColumn) {
		return new ExcelSquareRange(this, beginRow, endRow, beginColumn, endColumn);
	}
	
	public ExcelHorizontalRange horizontalRange(int row, int beginColumn, int endColumn) {
		return new ExcelHorizontalRange(this, row, beginColumn, endColumn);
	}
	
	public ExcelVerticalRange verticalRange(int column, int beginRow, int endRow) {
		return new ExcelVerticalRange(this, column, beginRow, endRow);
	}
}
