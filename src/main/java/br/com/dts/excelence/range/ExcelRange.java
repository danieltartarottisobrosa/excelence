package br.com.dts.excelence.range;

import br.com.dts.excelence.style.ExcelStyle;

public interface ExcelRange {
	public ExcelRange value(Object value);
	public ExcelRange style(ExcelStyle ...styles);
}
