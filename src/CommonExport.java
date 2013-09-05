import java.io.*;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class CommonExport {
	
	private Sheet colorSheet;
	
	public CommonExport() {
	}
	
	public CommonExport(Sheet colorSheet) {
		
		this.colorSheet = colorSheet;
	}

	public boolean referenceFont(String givenFont) {
		
		try {
			
			@SuppressWarnings("resource")
			Scanner scanFont = new Scanner(new FileReader("src/CartoCSS Fonts"));
			scanFont.useDelimiter(System.getProperty("line.separator"));
			
			while (scanFont.hasNext()) {  
				
				if(givenFont.trim().equalsIgnoreCase(scanFont.next())){
					return true;
				}
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		return false;
	}

	public String referenceColor(String givenColor) {
		
		for (int rowIndex = 4; rowIndex <= colorSheet.getLastRowNum(); rowIndex++) {
			
			Row tempRow = colorSheet.getRow(rowIndex);
			
			if(tempRow.getCell(0).toString().equalsIgnoreCase(givenColor)) {
				return tempRow.getCell(1).toString();
			}
		}
		return "";
	}
	
	public void appendLayerConditions(Row row, String storeCartoCSS, BufferedWriter writer, String currentClass) throws IOException {
		
		Double tempDouble;
		storeCartoCSS = storeCartoCSS.concat("#"+ currentClass);

		if (row.getCell(4) != null && !row.getCell(4).toString().equalsIgnoreCase("")) {
			storeCartoCSS = storeCartoCSS.concat("["+ row.getCell(4) + "]");
		}

		if (row.getCell(7) != null && !row.getCell(7).toString().equalsIgnoreCase("")) {
			tempDouble = new Double(row.getCell(7).toString());
			storeCartoCSS = storeCartoCSS.concat("[zoom>"+ (int)Math.round(tempDouble-1) + "]");
		}

		if (row.getCell(8) != null && !row.getCell(8).toString().equalsIgnoreCase("")) {
			tempDouble = new Double(row.getCell(8).toString());
			storeCartoCSS = storeCartoCSS.concat("[zoom<"+ (int)Math.round(tempDouble+1) + "]");
		}
		writer.append(storeCartoCSS);
	}
}
