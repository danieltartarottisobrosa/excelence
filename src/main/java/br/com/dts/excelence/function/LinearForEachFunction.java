package br.com.dts.excelence.function;

import br.com.dts.excelence.range.ExcelRange;

public interface LinearForEachFunction<R extends ExcelRange> {

	public void apply(R range, int i);
}
