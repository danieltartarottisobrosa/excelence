package br.com.dts.excelence;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import br.com.dts.excelence.exception.FileExtensionNotSupportedException;

public class ExcelWorkbook {
	private final Workbook poiWorkbook;
	private Path filePath;
	
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
}
