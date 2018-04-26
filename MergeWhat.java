import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MergeWhat {
	static int heading = 0;
	static int a = 0;
	static int top = 0;
	static int rows = 0;
	static int totalRows = 0;
	static ArrayList<String> names = new ArrayList<String>();
	static File log;
	static int count = 0;

	public static void main(String[] args) throws IOException {
		String folder = args[0];
		File file = new File(folder);
		File[] arr = file.listFiles();
		List<FileInputStream> list = new ArrayList<FileInputStream>();
		for (int i = 0; i < arr.length; i++) {
			names.add(arr[i].getName());
			FileInputStream inputStream = new FileInputStream(new File(folder + "/" + arr[i].getName()));
			list.add(inputStream);
		}
		File total2 = new File(folder + "/CurrentEmployees.xlsx");
		File total = new File(folder + "/NewHires.xlsx");
		log = new File(folder + "/log.csv");
		BufferedWriter bw = new BufferedWriter(new FileWriter(log.getAbsolutePath()));
		bw.write("File Name, Rows");
		bw.newLine();
		a++;
		mergeExcelFiles(total, list, bw);
		a--;
		list.clear();
		heading = 0;
		for (int i = 0; i < arr.length; i++) {
			FileInputStream inputStream = new FileInputStream(
					new File(folder + "/" + arr[i].getName()));
			list.add(inputStream);
		}
		
		bw.write("Number of New Hires," + totalRows);
		bw.newLine();
		bw.newLine();
		bw.write("File Name,Number of Rows");
		bw.newLine();
		bw.flush();
		totalRows = 0;
		mergeExcelFiles(total2, list, bw);
		bw.write("Number of Current Employees," + totalRows);
		bw.close();
	}

	public static void mergeExcelFiles(File file, List<FileInputStream> list, BufferedWriter bw) throws IOException {
		count++;
		XSSFWorkbook book = new XSSFWorkbook();
		XSSFSheet sheet = book.createSheet(file.getName());

		if (a == 0) {
			int i = 0;
			for (FileInputStream fin : list) {
				i++;
				System.out.println("Creating workbook");
				XSSFWorkbook b = new XSSFWorkbook(fin);
				System.out.println("Created workbook");
				copySheets(book, sheet, b.getSheetAt(0), 2);
				if(i == 1) {
					rows -= 2;
				}
				bw.write("\"" + names.get(i - 1) + "\"" + "," + rows);
				System.out.println(names.get(i - 1) + "was merged.");
				totalRows += rows;
				bw.newLine();
				bw.flush();
				b.close();
			}
		} else {
			int i = 0;
			for (FileInputStream fin : list) {
				i++;
				System.out.println("Creating workbook");
				XSSFWorkbook b = new XSSFWorkbook(fin);
				System.out.println("Created workbook");
				copySheets(book, sheet, b.getSheetAt(1), 1);
				if(i == 1) {
					rows -= 2;
				}
				bw.write("\"" + names.get(i - 1) + "\"" + "," + rows);
				System.out.println(names.get(i - 1) + "was merged.");
				totalRows += rows;
				bw.newLine();
				bw.flush();
				b.close();
			}
		}


		try {
			writeFile(book, file);
		} catch (Exception e) {
			e.printStackTrace();
		}
		bw.flush();
	}

	protected static void writeFile(XSSFWorkbook book, File file) throws Exception {
		FileOutputStream out = new FileOutputStream(file);
		book.write(out);
		out.close();
	}

	private static void copySheets(XSSFWorkbook newWorkbook, XSSFSheet newSheet, XSSFSheet sheet, int fileType) {
		copySheets(newWorkbook, newSheet, sheet, true, fileType);
	}

	private static void copySheets(XSSFWorkbook newWorkbook, XSSFSheet newSheet, XSSFSheet sheet, boolean copyStyle, int fileType) {
		rows = 0;
		int newRownumber = newSheet.getLastRowNum();
		int maxColumnNum = 0;
		Map<Integer, XSSFCellStyle> styleMap = (copyStyle) ? new HashMap<Integer, XSSFCellStyle>() : null;
		int number = 0;
		if (heading == 0) {
			number = sheet.getFirstRowNum();
			heading++;
			top--;
		} else {
			number = sheet.getFirstRowNum() + 2;
			newRownumber = newSheet.getLastRowNum();
			if(a == 0) {
				newRownumber -= 1;
			}
			else {
				newRownumber -= 2;
			}
		}
		for (int i = number; i <= sheet.getLastRowNum(); i++) {
			XSSFRow srcRow = sheet.getRow(i);
			XSSFRow destRow = newSheet.createRow(i + newRownumber);
			if (srcRow != null && !isRowEmpty(srcRow, fileType)) {
				copyRow(newWorkbook, sheet, newSheet, srcRow, destRow, styleMap);
				if (srcRow.getLastCellNum() > maxColumnNum) {
					maxColumnNum = srcRow.getLastCellNum();
				}
			} else {
				break;
			}
		}
		for (int i = 0; i <= maxColumnNum; i++) {
			newSheet.setColumnWidth(i, sheet.getColumnWidth(i));
		}
	}

	public static void copyRow(XSSFWorkbook newWorkbook, XSSFSheet srcSheet, XSSFSheet destSheet, XSSFRow srcRow,
			XSSFRow destRow, Map<Integer, XSSFCellStyle> styleMap) {
		destRow.setHeight(srcRow.getHeight());

		for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {
			if (j == -1) {
				return;
			}
			XSSFCell oldCell = srcRow.getCell(j);
			XSSFCell newCell = destRow.getCell(j);
			if (oldCell != null) {
				if (newCell == null) {
					newCell = destRow.createCell(j);
				}
				copyCell(newWorkbook, oldCell, newCell, styleMap, j);
			}
		}
		rows++;
	}

	public static void copyCell(XSSFWorkbook newWorkbook, XSSFCell oldCell, XSSFCell newCell,
			Map<Integer, XSSFCellStyle> styleMap, int num) {
		if (styleMap != null) {
			int stHashCode = oldCell.getCellStyle().hashCode();
			XSSFCellStyle newCellStyle = styleMap.get(stHashCode);
			if (newCellStyle == null) {
				newCellStyle = newWorkbook.createCellStyle();
				newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
				styleMap.put(stHashCode, newCellStyle);
			}
			newCell.setCellStyle(newCellStyle);
		}
		switch (oldCell.getCellType()) {
		case XSSFCell.CELL_TYPE_STRING:
			newCell.setCellValue(oldCell.getRichStringCellValue());
			break;
		case XSSFCell.CELL_TYPE_NUMERIC:
			newCell.setCellValue(oldCell.getNumericCellValue());
			break;
		case XSSFCell.CELL_TYPE_BLANK:
			newCell.setCellType(XSSFCell.CELL_TYPE_BLANK);
			break;
		case XSSFCell.CELL_TYPE_BOOLEAN:
			newCell.setCellValue(oldCell.getRichStringCellValue());
			break;
		case XSSFCell.CELL_TYPE_ERROR:
			//newCell.setCellErrorValue(oldCell.getErrorCellValue());
			newCell.setCellValue(oldCell.getErrorCellString());
			break;
		case XSSFCell.CELL_TYPE_FORMULA:
			switch(oldCell.getCachedFormulaResultType()) {
			case XSSFCell.CELL_TYPE_STRING:
				newCell.setCellValue(oldCell.getRichStringCellValue());
				break;
			case XSSFCell.CELL_TYPE_NUMERIC:
				newCell.setCellValue(oldCell.getNumericCellValue());
				break;
			}
			break;
		default:
			break;
		}
	}

	public static boolean isRowEmpty(XSSFRow row, int fileType) {
		
		if (row.getFirstCellNum() > -1) {
			for (int c = row.getFirstCellNum(); c < fileType; c++) {
				XSSFCell cell = row.getCell(c);
				if (cell != null && cell.getCellType() != XSSFCell.CELL_TYPE_BLANK)
					return false;
			}
			return true;
		} else {
			return false;
		}
	}
}
