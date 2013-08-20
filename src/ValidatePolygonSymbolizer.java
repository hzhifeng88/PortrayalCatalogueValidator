import java.util.ArrayList;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ValidatePolygonSymbolizer extends CommonValidator{

	private Sheet polygonSheet;

	public ValidatePolygonSymbolizer(Sheet sheet, ArrayList<String> list , HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet, list, kit, doc);
		this.polygonSheet = sheet;
		
		checkMergedCells();
		checkEmptyRows();

		if (printFormatError() == false) {
			checkStyleID();
			checkColorValid();
			checkMissingAttributes();
			printValueError();
		}
	}

	public void checkStyleID() {

		for (int rowIndex = 4; rowIndex <= polygonSheet.getLastRowNum(); rowIndex++) {

			checkIDAndDuplicate('A', "F", rowIndex, 5);
		}
	}
	
	// Column K and Q store color values or references
	public void checkColorValid() {

		int rowIndex;

		for (rowIndex = 4; rowIndex <= polygonSheet.getLastRowNum(); rowIndex++) {

			Row row = polygonSheet.getRow(rowIndex);
			matchColor(row.getCell(10).toString(), "K", rowIndex);
			matchColor(row.getCell(16).toString(), "Q", rowIndex);
		}
	}
}
