package reading.files;

import java.util.ArrayList;
import java.util.Map;

public class excelFileTestCases {

	public static void main(String[] args) {
		ReadExcelFile file = new ReadExcelFile();
		Map<String, ArrayList<Object>> readData = file.readWBColumnWise("testSample.xlsx");
		for (Map.Entry<String, ArrayList<Object>> map:readData.entrySet()) {
			System.out.print(map.getKey()+" : ");
			for (Object obj:map.getValue()) {
				System.out.print(obj+",");
			}
			System.out.println();
		}

	}

}
