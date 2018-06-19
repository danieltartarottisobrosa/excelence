package br.com.dts.excelence;

import java.io.IOException;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.mockito.runners.MockitoJUnitRunner;

import br.com.dts.excelence.range.ExcelCell;
import br.com.dts.excelence.range.ExcelHorizontalRange;
import br.com.dts.excelence.range.ExcelLinearRange;
import br.com.dts.excelence.range.ExcelRange;
import br.com.dts.excelence.range.ExcelVerticalRange;
import br.com.dts.excelence.style.builder.ExcelBorder;
import br.com.dts.excelence.style.builder.ExcelFont;
import br.com.dts.excelence.style.builder.ExcelFormat;

@RunWith(MockitoJUnitRunner.class)
public class GeneralUseTest {

	@Test
	@SuppressWarnings("unused")
	public void test1() throws IOException {
		String filePath = getResourcesPath("test1.xlsx");
		
		ExcelWorkbook workbook = Excelence.createWorkbook();
		ExcelWorkbook workbook2 = Excelence.openWorkbook(filePath);
		
		ExcelSheet sheet = workbook.createSheet("Sheet 1");
		ExcelCell cell1 = sheet.cell(0, 0).value("Cell value");
		String value = cell1.value(String.class);
		String valueNotFound = sheet.cell(0, 1).valueDefault("Not found!");
		
		ExcelLinearRange range1 = sheet.horizontalRange(0, 0, 4);
		ExcelRange range2 = range1;
		ExcelRange range3 = cell1;
		
		sheet.range(0, 0, 10, 10)
			.style(
				new ExcelBorder().thickness(10).color("blue").top(true).bottom(true),
				new ExcelFont().size(14).color("red"),
				new ExcelFormat().format("#,##0.00"))
			.value(1234.567);
		
		ExcelHorizontalRange columns = sheet.horizontalRange(5, 0, 2);
		columns.values("Name", "Age", "Phone");
		
		columns.forEach((cell, col) -> {
			System.out.println(String.format("%d) %s", col, cell.value(String.class)));
		});
		
		ExcelVerticalRange rows = sheet.verticalRange(5, 0, 10);
		rows.style(new ExcelBorder().all(true));
		
		workbook.save(filePath);
	}
	
	private String getResourcesPath(String resource) {
		return getClass().getClassLoader().getResource(resource).getPath();
	}
}
