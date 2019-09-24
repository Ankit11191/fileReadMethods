package reading.files;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;;

public class ReadFile {

	private Workbook creataWorkBook(String fileName) {

		Workbook wb = null;
		try {
			FileInputStream inputFilepath = new FileInputStream(new File(
					System.getProperty("user.dir") + File.separator + "resources" + File.separator + fileName));
			String fileExtensionName = fileName.substring(fileName.indexOf("."));

			if (fileExtensionName.equals(".xlsx")) {
				wb = new XSSFWorkbook(inputFilepath);
			} else if (fileExtensionName.equals(".xls")) {
				wb = new HSSFWorkbook(inputFilepath);
			}
		} catch (IOException e) {
			System.out.println(e.getLocalizedMessage());
		}
		return wb;
	}

	public ArrayList<ArrayList<Object>> readWorkBook(String fileName) {
		ArrayList<ArrayList<Object>> alMain = new ArrayList<ArrayList<Object>>();

		ReadFile file = new ReadFile();
		Workbook getWB = file.creataWorkBook(fileName);
		Sheet sheet = getWB.getSheetAt(0);
		Iterator<Row> rowItr = sheet.rowIterator();
		FormulaEvaluator evaluator = getWB.getCreationHelper().createFormulaEvaluator();

		while (rowItr.hasNext()) {
			ArrayList<Object> alInner = new ArrayList<Object>();
			Row row = rowItr.next();
			Iterator<Cell> cellItr = row.cellIterator();
			while (cellItr.hasNext()) {
				Cell cell = cellItr.next();
				switch (cell.getCellType()) {
				case NUMERIC:
					alInner.add(cell.getNumericCellValue());
					break;
				case STRING:
					alInner.add(cell.getStringCellValue());
					break;
				case BLANK:
					alInner.add(null);
					break;
				case BOOLEAN:
					alInner.add(cell.getBooleanCellValue());
					break;
				case FORMULA:
					alInner.add(evaluator.evaluateFormulaCell(cell));
					break;
				case ERROR:
					alInner.add(cell.getErrorCellValue());
					break;
				case _NONE:
					alInner.add(cell.getDateCellValue());
					break;
				default:
					break;
				}
			}
			alMain.add(alInner);
		}
		return alMain;
	}
}
