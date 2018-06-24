package br.com.dts.excelence.valueparser;

import org.apache.poi.ss.usermodel.Cell;

public class IntegerParser implements ValueParser<Integer> {
	private DoubleParser parser = new DoubleParser();
	
	@Override
	public Integer readValue(Cell poiCell) {
		return parser.readValue(poiCell).intValue();
	}

	@Override
	public void writeValue(Cell poiCell, Object value) {
		parser.writeValue(poiCell, value);
	}

}
