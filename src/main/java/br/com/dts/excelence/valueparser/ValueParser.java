package br.com.dts.excelence.valueparser;

import org.apache.poi.ss.usermodel.Cell;

public interface ValueParser<T> {

	T readValue(Cell poiCell);
	void writeValue(Cell poiCell, Object value);
	
}
