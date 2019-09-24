package reading.files;

import java.util.ArrayList;

public class excelFileTestCases {

	public static void main(String[] args) {
		ReadFile file = new ReadFile();
		ArrayList<ArrayList<Object>> SheetData = file.readWorkBook("testSample.xlsx");
		for (ArrayList<Object> al : SheetData) {
			for (Object obj : al) {
				System.out.print(obj + "\t");
			}
			System.out.println();
		}

	}

}
