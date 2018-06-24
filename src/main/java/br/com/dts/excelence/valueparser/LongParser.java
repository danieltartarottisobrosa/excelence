package br.com.dts.excelence.valueparser;

import org.apache.poi.ss.usermodel.Cell;

public class LongParser implements ValueParser<Long> {
	private DoubleParser parser = new DoubleParser();
	
	@Override
	public Long readValue(Cell poiCell) {
		return parser.readValue(poiCell).longValue();
	}

	@Override
	public void writeValue(Cell poiCell, Object value) {
		parser.writeValue(poiCell, value);
	}

}
