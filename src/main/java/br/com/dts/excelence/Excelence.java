package br.com.dts.excelence;

import java.io.IOException;

public final class Excelence {
	
	public Excelence() {}
	
	public static ExcelWorkbook createWorkbook() {
		return new ExcelWorkbook();
	}
	
	public static ExcelWorkbook openWorkbook(String filePath) throws IOException {
		return new ExcelWorkbook(filePath);
	}
}
