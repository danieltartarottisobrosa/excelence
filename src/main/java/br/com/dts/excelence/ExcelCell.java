package br.com.dts.excelence;

import java.util.Objects;
import java.util.Optional;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;

import br.com.dts.excelence.style.ExcelStyle;

public class ExcelCell {
	private ExcelSheet sheet;
	private final Cell poiCell;
	
	//
	//  Constructors
	//
	
	public ExcelCell(ExcelSheet sheet, int row, int col) {
		this.sheet = sheet;
		
		Sheet poiSheet = sheet.poiSheet();
		
		if (Objects.isNull(poiSheet.getRow(row))) poiSheet.createRow(row);
		poiCell = poiSheet.getRow(row).getCell(col, MissingCellPolicy.CREATE_NULL_AS_BLANK);
	}
	
	//
	//  Getters and Setters
	//

	public ExcelSheet sheet() {
		return sheet;
	}
	
	public int row() {
		return poiCell.getRowIndex();
	}
	
	public int col() {
		return poiCell.getColumnIndex();
	}
	
	public Cell poiCell() {
		return poiCell;
	}
	
	//
	//  Public Methods
	//

	@SuppressWarnings("unchecked")
	public <T extends Object> T valueDefault(T defaultValue) {
		T value = (T) value(defaultValue.getClass());
		return Optional.ofNullable(value).orElse(defaultValue);
	}
	
	@SuppressWarnings("unchecked")
	public <T extends Object> T value(Class<T> type) {
		return (T) sheet.workbook().valueParser(type).readValue(poiCell);
	}
	
	public ExcelCell value(Object value) {
		Object val = Optional.ofNullable(value).orElse("");
		sheet.workbook().valueParser(value.getClass()).writeValue(poiCell, val);
		return this;
	}

	public <T extends Object> ExcelCell valueDefault(T value, T defaultValue) {
		T val = Optional.ofNullable(value).orElse(defaultValue);
		return value(val);
	}

	public ExcelCell style(ExcelStyle... styles) {

		return this;
	}

}
