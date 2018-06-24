package br.com.dts.excelence.valueparser;

import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;

public class DateParser implements ValueParser<Date> {

	@Override
	public Date readValue(Cell poiCell) {
		return poiCell.getDateCellValue();
	}

	@Override
	public void writeValue(Cell poiCell, Object value) {
		poiCell.setCellValue((Date) value);
	}

}
