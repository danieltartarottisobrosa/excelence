package br.com.dts.excelence.range;

import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Stream;

import org.apache.poi.ss.util.CellRangeAddress;

import br.com.dts.excelence.ExcelSheet;
import br.com.dts.excelence.style.ExcelStyle;

public class ExcelSquareRange implements ExcelRange {
	private ExcelSheet sheet;
	private CellRangeAddress poiRange;
	
	//
	//  Constructors
	//
	
	public ExcelSquareRange(ExcelSheet sheet, int beginRow, int beginColumn, int endRow, int endColumn) {
		this.sheet = sheet;
		poiRange = new CellRangeAddress(beginRow, endRow, beginColumn, endColumn);
	}
	
	//
	//  Getters and Setters
	//
	
	public ExcelSheet sheet() {
		return sheet;
	}
	
	public final int beginRow() {
		return poiRange.getFirstRow();
	}
	
	public final int endRow() {
		return poiRange.getLastRow();
	}
	
	public final int beginCol() {
		return poiRange.getFirstColumn();
	}
	
	public final int endCol() {
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
	
	public ExcelCell cell(int row, int col) {
		return new ExcelCell(sheet, row, col);
	}
	
	public ExcelSquareRange value(Object value) {
		stream().forEach(cell -> cell.value(value));
		return this;
	}

	public ExcelSquareRange style(ExcelStyle... styles) {

		return this;
	}
	
	public Stream<ExcelCell> stream() {
		AtomicInteger row = new AtomicInteger(beginRow());
		AtomicInteger col = new AtomicInteger(beginCol());
		
		int qtdRows = endRow() - beginRow() + 1;
		int qtdCols = endCol() - beginCol() + 1;
		int qtdCells = qtdRows * qtdCols;
		
		return Stream.generate(() -> {
			if (col.get() > endCol()) {
				row.incrementAndGet();
				col.set(0);
			}
			return cell(row.get(), col.getAndIncrement());
		}).limit(qtdCells);
	}
	
	public Stream<ExcelCell> colStream(int col) {
		AtomicInteger row = new AtomicInteger(beginRow());
		return Stream.generate(() -> cell(col, row.get()));
	}
	
	public Stream<ExcelCell> rowStream(int row) {
		AtomicInteger col = new AtomicInteger(beginCol());
		return Stream.generate(() -> cell(row, col.get()));
	}
	
	public Stream<ExcelHorizontalRange> horizontalStream() {
		AtomicInteger row = new AtomicInteger(beginRow());
		return Stream.generate(() -> new ExcelHorizontalRange(sheet, row.get(), beginCol(), endCol()));
	}
	
	public Stream<ExcelVerticalRange> verticalStream() {
		AtomicInteger row = new AtomicInteger(beginRow());
		return Stream.generate(() -> new ExcelVerticalRange(sheet, row.get(), beginCol(), endCol()));
	}

}
