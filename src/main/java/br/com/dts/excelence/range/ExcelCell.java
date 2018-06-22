package br.com.dts.excelence.range;

import java.util.Date;
import java.util.Objects;
import java.util.Optional;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;

import br.com.dts.excelence.ExcelSheet;
import br.com.dts.excelence.style.ExcelStyle;

public class ExcelCell implements ExcelRange {
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
		if (type.isAssignableFrom(Boolean.class))
			return (T) (Boolean) poiCell.getBooleanCellValue();
		
		if (type.isAssignableFrom(Date.class))
			return (T) poiCell.getDateCellValue();
		
		if (type.isAssignableFrom(Double.class))
			return (T) (Double) poiCell.getNumericCellValue();
			
		if (type.isAssignableFrom(Integer.class))
			return (T) (Integer) ((Double) poiCell.getNumericCellValue()).intValue();
			
		if (type.isAssignableFrom(Long.class))
			return (T) (Long) ((Double) poiCell.getNumericCellValue()).longValue();
		
		if (type.isAssignableFrom(String.class)) {
			return (T) poiCell.getStringCellValue();
		}
		 
		return null;
	}
	
	public ExcelCell value(Object value) {
		if (value == null)
			poiCell.setCellValue("");
		else if (value instanceof Boolean)
			poiCell.setCellValue((Boolean) value);
		else if (value instanceof Date)
			poiCell.setCellValue((Date) value);
		else if (value instanceof Double)
			poiCell.setCellValue((Double) value);
		else if (value instanceof Integer)
			poiCell.setCellValue(((Integer) value).doubleValue());
		else if (value instanceof Long)
			poiCell.setCellValue(((Long) value).doubleValue());
		else
			poiCell.setCellValue(value.toString());
		
		return this;
	}

	public <T extends Object> ExcelCell valueDefault(T value, T defaultValue) {
		T val = Optional.ofNullable(value).orElse(defaultValue);
		return value(val);
	}

	public ExcelRange style(ExcelStyle... styles) {

		return this;
	}

}
