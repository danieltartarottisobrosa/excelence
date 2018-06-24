package br.com.dts.excelence.range;

import java.util.stream.Stream;

import br.com.dts.excelence.ExcelCell;
import br.com.dts.excelence.style.ExcelStyle;

public interface ExcelRange {
	public ExcelRange value(Object value);
	public ExcelRange style(ExcelStyle ...styles);
	public Stream<ExcelCell> stream();
}
