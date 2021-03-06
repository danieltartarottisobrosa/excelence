package br.com.dts.excelence;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.junit.Assert;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.mockito.runners.MockitoJUnitRunner;

import br.com.dts.excelence.range.ExcelHorizontalRange;
import br.com.dts.excelence.range.ExcelLinearRange;
import br.com.dts.excelence.range.ExcelRange;
import br.com.dts.excelence.range.ExcelSquareRange;
import br.com.dts.excelence.range.ExcelVerticalRange;
import br.com.dts.excelence.style.ExcelBorder;
import br.com.dts.excelence.style.ExcelFont;
import br.com.dts.excelence.style.ExcelFormat;
import br.com.dts.excelence.valueparser.ValueParser;

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
		
		ExcelSquareRange square = sheet.range(0, 0, 9, 9)
			.style(
				new ExcelBorder().thickness(10).color("blue").top(true).bottom(true),
				new ExcelFont().size(14).color("red"),
				new ExcelFormat().format("#,##0.00"))
			.value(1234.567);
			
		square.stream().map(c -> c.value(Double.class)).forEach(System.out::print);
		
		ExcelHorizontalRange columns = sheet.horizontalRange(5, 0, 2);
		columns.values("Name", "Age", "Phone");
		columns.stream().map(c -> c.value(String.class)).forEach(System.out::print);
		
		ExcelVerticalRange rows = sheet.verticalRange(5, 0, 4);
		rows.style(new ExcelBorder().all(true));
		rows.values("A", "B", "C", "D", "E");
		
		rows.stream().map(c -> c.value(String.class)).forEach(System.out::println);
		
		workbook.save(filePath);
	}
	
	@Test
	public void readAndWriteStringCellValue() {
		ExcelWorkbook workbook = new ExcelWorkbook();
		ExcelSheet sheet = workbook.createSheet("Sheet 1");
		ExcelCell cell = sheet.cell(0, 0).value("Some value");
		String value = cell.value(String.class);
		Assert.assertEquals("Some value", value);
		
		value = sheet.cell(0, 0).value(String.class);
		Assert.assertEquals("Some value", value);
	}
	
	@Test
	public void customValueParser() {
		ExcelWorkbook workbook = Excelence.createWorkbook();
		workbook.valueParser(ComplexType.class, new ComplexTypeParser());
		
		ExcelSheet sheet = workbook.createSheet("Sheet");
		
		ExcelCell cell = sheet.cell(3, 3).value(new ComplexType("Daniel"));
		System.out.println(cell.value(ComplexType.class));
	}
	
	private String getResourcesPath(String resource) {
		return getClass().getClassLoader().getResource(resource).getPath();
	}
	
	static class ComplexType {
		private final String name;
		
		public ComplexType(String name) {
			this.name = name;
		}
		
		public String getName() {
			return name;
		}
	}
	
	static class ComplexTypeParser implements ValueParser<ComplexType> {
		@Override
		public ComplexType readValue(Cell poiCell) {
			return new ComplexType(poiCell.getStringCellValue());
		}
		@Override
		public void writeValue(Cell poiCell, Object value) {
			poiCell.setCellValue(((ComplexType) value).getName());
		}
	}
}

