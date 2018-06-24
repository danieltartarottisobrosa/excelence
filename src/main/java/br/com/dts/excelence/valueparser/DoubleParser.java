package br.com.dts.excelence.valueparser;

import org.apache.poi.ss.usermodel.Cell;

public class DoubleParser implements ValueParser<Double> {

	@Override
	public Double readValue(Cell poiCell) {
		return poiCell.getNumericCellValue();
	}

	@Override
	public void writeValue(Cell poiCell, Object value) {
		poiCell.setCellValue((double) value);
	}

}
