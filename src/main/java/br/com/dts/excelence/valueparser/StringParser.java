package br.com.dts.excelence.valueparser;

import org.apache.poi.ss.usermodel.Cell;

public class StringParser implements ValueParser<String> {

	@Override
	public String readValue(Cell poiCell) {
		return poiCell.getStringCellValue();
	}

	@Override
	public void writeValue(Cell poiCell, Object value) {
		poiCell.setCellValue(value.toString());
	}

}
