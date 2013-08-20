import java.util.ArrayList;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ValidateTextSymbolizer extends CommonValidator {

	private Sheet textSheet;

	public ValidateTextSymbolizer(Sheet sheet, ArrayList<String> list,  HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet, list, kit, doc);
		this.textSheet = sheet;

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

		for (int rowIndex = 4; rowIndex <= textSheet.getLastRowNum(); rowIndex++) {
			
			checkIDAndDuplicate('T', "F", rowIndex, 5);
		}
	}

	// Column T store color values or references
	public void checkColorValid() {

		int rowIndex;

		for (rowIndex = 4; rowIndex <= textSheet.getLastRowNum(); rowIndex++) {

			Row row = textSheet.getRow(rowIndex);
			matchColor(row.getCell(19).toString(),"T", rowIndex);

		}
	}
}
