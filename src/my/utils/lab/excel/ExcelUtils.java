package my.utils.lab.excel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

public class ExcelUtils<T> {

	public Class<T> clazz;

	public ExcelUtils(Class<T> clazz) {
		this.clazz = clazz;
	}

	public boolean exportExcel(List<T> lists[], String sheetNames[], OutputStream outputStream) {

		if (lists.length != sheetNames.length) {
			return false;
		}

		// Declare a workbook
		SXSSFWorkbook workbook = new SXSSFWorkbook();

		for (int i = 0; i < lists.length; i++) {
			List<T> list = lists[i];
			String sheetName = sheetNames[i];

			List<Field> fieldList = getMappedColumns(clazz, null);

			// Declare a worksheet
			SXSSFSheet sheet = workbook.createSheet(sheetName);

			workbook.setSheetName(i, sheetName);

			// Declare a empty row
			SXSSFRow row = null;
			// Declare a empty cell in row
			SXSSFCell cell = null;
			// Create a row
			row = sheet.createRow(0);

			// Create headers
			for (int ij = 0; ij < fieldList.size(); ij++) {
				// Get field from list
				Field field = fieldList.get(ij);
				// Get annotation from field
				ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
				// Create cell in row
				cell = row.createCell(ij);
				// Set cell type to string
				cell.setCellType(CellType.STRING);
				// Set cell value to header name
				cell.setCellValue(excelColumn.name());
			}

			// Declare index of first row exclude header
			int start = 0;
			// Declare index of last row
			int end = list.size();

			// Write record in a row
			for (int ik = start; ik < end; ik++) {
				// Calculate index of current row
				row = sheet.createRow(ik + 1 - start);
				// Get record from the list
				T vo = list.get(ik);

				for (int iki = 0; iki < fieldList.size(); iki++) {
					Field field = fieldList.get(iki);
					// Access private fields
					field.setAccessible(true);

					try {
						cell = row.createCell(iki);
						cell.setCellType(CellType.STRING);
						cell.setCellValue(field.get(vo) == null ? "" : String.valueOf(field.get(vo)));
					} catch (IllegalAccessException e) {
						e.printStackTrace();
					}

				}

			}

		}

		try {
			outputStream.flush();
			workbook.write(outputStream);
			outputStream.close();

			return true;
		} catch (IOException e) {
			e.printStackTrace();

			return false;
		}
	}

	public boolean exportExcel(List<T> list, String sheetName, OutputStream output) {
		ArrayList[] lists = new ArrayList[1];
		lists[0] = (ArrayList) list;

		String[] sheetNames = new String[1];
		sheetNames[0] = sheetName;

		return exportExcel(lists, sheetNames, output);
	}

	private List<Field> getMappedColumns(Class clazz, List<Field> fields) {
		if (fields == null) {
			fields = new ArrayList<>();
		}

		// Get all field inside excel view object class
		Field[] allFields = clazz.getDeclaredFields();

		// Get all fields include @ExcelColumn annotation and put it into the list
		for (Field field : allFields) {
			if (field.isAnnotationPresent(ExcelColumn.class)) {
				fields.add(field);
			}
		}

		return fields;
	}
}
