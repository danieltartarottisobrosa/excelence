package br.com.dts.excelence.valueparser;

import org.apache.poi.ss.usermodel.Cell;

public class BooleanParser implements ValueParser<Boolean> {

	@Override
	public Boolean readValue(Cell poiCell) {
		return poiCell.getBooleanCellValue();
	}

	@Override
	public void writeValue(Cell poiCell, Object value) {
		poiCell.setCellValue((Boolean) value);
	}

}
