package br.com.dts.excelence;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import br.com.dts.excelence.exception.FileExtensionNotSupportedException;
import br.com.dts.excelence.exception.ParserNotFoundException;
import br.com.dts.excelence.valueparser.BooleanParser;
import br.com.dts.excelence.valueparser.DateParser;
import br.com.dts.excelence.valueparser.DoubleParser;
import br.com.dts.excelence.valueparser.IntegerParser;
import br.com.dts.excelence.valueparser.LongParser;
import br.com.dts.excelence.valueparser.StringParser;
import br.com.dts.excelence.valueparser.ValueParser;

public class ExcelWorkbook {
	private final Workbook poiWorkbook;
	private Path filePath;
	private final Map<Class<?>, ValueParser<? extends Object>> valueParsers = createDefaultParsers();
	
	//
	//  Constructors
	//
	
	public ExcelWorkbook() {
		poiWorkbook = new XSSFWorkbook();
	}

	public ExcelWorkbook(String filePath) throws IOException {
		this(Paths.get(filePath));
	}
	
	public ExcelWorkbook(Path filePath) throws IOException {
		this.filePath = filePath;
		
		final String fileName = this.filePath.getFileName().toString().toLowerCase();
		
		if (fileName.endsWith(".xlsx")) {
			poiWorkbook = new XSSFWorkbook(filePath.toString());
			
		} else if (fileName.endsWith(".xls")) {
			InputStream stream = new FileInputStream(filePath.toString());
			poiWorkbook = new HSSFWorkbook(stream);
			stream.close();
			
		} else {
			poiWorkbook = null;
			throw new FileExtensionNotSupportedException();
		}
	}

	//
	//  Getters and Setters
	//
	
	public Path filePath() {
		return filePath;
	}
	
	public Workbook poiWorkbook() {
		return poiWorkbook;
	}
	
	//
	//  Public Methods
	//

	public ExcelSheet createSheet(String sheetName) {
		return new ExcelSheet(this, sheetName);
	}
	
	public ExcelWorkbook valueParser(Class<?> type, ValueParser<? extends Object> parser) {
		valueParsers.put(type, parser);
		return this;
	}
	
	public ValueParser<? extends Object> valueParser(Class<?> type) {
		return Optional.ofNullable(valueParsers.get(type))
				.orElseThrow(() -> new ParserNotFoundException(type));
	}

	public ExcelWorkbook save() throws IOException {
		return save(filePath);
	}
	
	public ExcelWorkbook save(String filePath) throws IOException {
		return save(Paths.get(filePath));
	}
	
	public ExcelWorkbook save(Path filePath) throws IOException {
		OutputStream stream = new FileOutputStream(filePath.toString());
		poiWorkbook.write(stream);
		stream.close();
		
		this.filePath = filePath;
		return this;
	}
	
	//
	//  Private Methods
	//
	
	private Map<Class<?>, ValueParser<? extends Object>> createDefaultParsers() {
		Map<Class<?>, ValueParser<? extends Object>> valueParsers = new HashMap<>();
		valueParsers.put(Boolean.class, new BooleanParser());
		valueParsers.put(Date.class, new DateParser());
		valueParsers.put(Double.class, new DoubleParser());
		valueParsers.put(Integer.class, new IntegerParser());
		valueParsers.put(Long.class, new LongParser());
		valueParsers.put(String.class, new StringParser());
		return valueParsers;
	}
}
