package br.com.dts.excelence.range;

import java.util.stream.IntStream;

import org.apache.poi.ss.util.CellRangeAddress;

import br.com.dts.excelence.ExcelSheet;
import br.com.dts.excelence.function.LinearForEachFunction;
import br.com.dts.excelence.function.SquareForEachFunction;
import br.com.dts.excelence.style.ExcelStyle;

public class ExcelSquareRange implements ExcelRange {
	private ExcelSheet sheet;
	private CellRangeAddress poiRange;
	
	//
	//  Constructors
	//
	
	public ExcelSquareRange(ExcelSheet sheet, int beginRow, int endRow, int beginColumn, int endColumn) {
		this.sheet = sheet;
		poiRange = new CellRangeAddress(beginRow, endRow, beginColumn, endColumn);
	}
	
	//
	//  Getters and Setters
	//
	
	public ExcelSheet sheet() {
		return sheet;
	}
	
	public int beginRow() {
		return poiRange.getFirstRow();
	}
	
	public int endRow() {
		return poiRange.getLastRow();
	}
	
	public int beginCol() {
		return poiRange.getFirstColumn();
	}
	
	public int endCol() {
		return poiRange.getLastColumn();
	}
	
	public CellRangeAddress poiRange() {
		return poiRange;
	}

	//
	//  Public Methods
	//
	
	public ExcelSquareRange value(Object ...value) {
		
		return this;
	}
	
	public ExcelCell cell(int row, int col, Object value) {
		return new ExcelCell(sheet, row, col);
	}
	
	public ExcelSquareRange value(Object value) {
		return forEach((cell, row, col) -> cell.value(value));
	}

	public ExcelSquareRange style(ExcelStyle... styles) {

		return this;
	}
	
	public ExcelSquareRange forEach(SquareForEachFunction f) {
		IntStream.rangeClosed(beginRow(), endRow()).forEach(row ->
				IntStream.rangeClosed(beginCol(), endCol()).forEach(col ->
						f.apply(new ExcelCell(sheet, row, col), row, col)));
		
		return this;
	}
	
	public ExcelSquareRange forEachRow(LinearForEachFunction<ExcelRange> f) {
		IntStream.rangeClosed(beginRow(), endRow()).forEach(row -> {
			ExcelRange range = new ExcelHorizontalRange(sheet, row, beginCol(), endCol());
			f.apply(range, row);
		});
		
		return this;
	}
	
	public ExcelSquareRange forEachCol(LinearForEachFunction<ExcelRange> f) {
		IntStream.rangeClosed(beginCol(), endCol()).forEach(col -> {
			ExcelRange range = new ExcelVerticalRange(sheet, col, beginRow(), endRow());
			f.apply(range, col);
		});
		
		return this;
	}

}
