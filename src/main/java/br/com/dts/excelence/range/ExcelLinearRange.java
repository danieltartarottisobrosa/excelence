package br.com.dts.excelence.range;

import java.util.stream.IntStream;

import br.com.dts.excelence.ExcelSheet;
import br.com.dts.excelence.function.LinearForEachFunction;

public abstract class ExcelLinearRange implements ExcelRange {
	private ExcelSheet sheet;
	private final int i;
	private final int beginJ;
	private final int endJ;

	//
	// Constructors
	//

	public ExcelLinearRange(ExcelSheet sheet, int i, int beginJ, int endJ) {
		this.sheet = sheet;
		this.i = i;
		this.beginJ = beginJ;
		this.endJ = endJ;
	}

	//
	// Getters and Setters
	//

	public ExcelSheet sheet() {
		return sheet;
	}

	protected int i() {
		return i;
	}

	protected int beginJ() {
		return beginJ;
	}

	protected int endJ() {
		return endJ;
	}
	
	public int length() {
		return endJ() - beginJ() + 1;
	}

	//
	// Public Methods
	//

	public ExcelLinearRange values(Object... values) {
		int end = Math.min(length(), values.length);
		
		IntStream.range(0, end).forEach(x -> {
			cell(x + beginJ()).value(values[x]);
		});
		
		return this;
	}
	
	public <T extends Object> T value(int j, Class<T> type) {
		return cell(j).value(type);
	}
	
	public ExcelLinearRange value(Object value) {
		return forEach((cell, col) -> cell.value(value));
	}
	
	public ExcelLinearRange forEach(LinearForEachFunction<ExcelCell> f) {
		IntStream.range(0, length()).forEach(x -> f.apply(cell(x + beginJ()), x));
		return this;
	}

	//
	// Abstract Methods
	//

	public abstract ExcelCell cell(int j);

}