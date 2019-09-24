package reading.files;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;;

public class ReadExcelFile {

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

	public ArrayList<ArrayList<Object>> readWB(String fileName) {
		ArrayList<ArrayList<Object>> alMain = new ArrayList<ArrayList<Object>>();

		ReadExcelFile file = new ReadExcelFile();
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
	
	public ArrayList<ArrayList<Object>> readWB(String fileName, String SheetName) {
		ArrayList<ArrayList<Object>> alMain = new ArrayList<ArrayList<Object>>();

		ReadExcelFile file = new ReadExcelFile();
		Workbook getWB = file.creataWorkBook(fileName);
		Sheet sheet = getWB.getSheet(SheetName);
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
	
	public Map<String, ArrayList<Object>> readWBColumnWise(String fileName) {
		Map<String, ArrayList<Object>> excelMap=new HashMap<String, ArrayList<Object>>();

		ReadExcelFile file = new ReadExcelFile();
		Workbook getWB = file.creataWorkBook(fileName);
		Sheet sheet = getWB.getSheetAt(0);



		FormulaEvaluator evaluator = getWB.getCreationHelper().createFormulaEvaluator();

		for(int i=0; i<sheet.getRow(0).getPhysicalNumberOfCells(); i++) {
			ArrayList<Object> alInner = new ArrayList<Object>();
			Iterator<Row> rowItr = sheet.rowIterator();
			Row row = rowItr.next();
			String colName=row.getCell(i).getStringCellValue();
			while (rowItr.hasNext()) {
				row = rowItr.next();
				switch (row.getCell(i).getCellType()) {
				case NUMERIC:
					alInner.add(row.getCell(i).getNumericCellValue());
					break;
				case STRING:
					alInner.add(row.getCell(i).getStringCellValue());
					break;
				case BLANK:
					alInner.add(null);
					break;
				case BOOLEAN:
					alInner.add(row.getCell(i).getBooleanCellValue());
					break;
				case FORMULA:
					alInner.add(evaluator.evaluateFormulaCell(row.getCell(i)));
					break;
				case ERROR:
					alInner.add(row.getCell(i).getErrorCellValue());
					break;
				case _NONE:
					alInner.add(row.getCell(i).getDateCellValue());
					break;
				default:
					break;
				}
			}
			excelMap.put(colName,alInner);
		}
		return excelMap;
	}
}
