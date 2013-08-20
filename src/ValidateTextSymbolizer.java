import java.util.ArrayList;

import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ValidateTextSymbolizer extends CommonValidator {

	private Sheet textSheet;

	public ValidateTextSymbolizer(Sheet sheet, Workbook originalWorkbook, ArrayList<String> list,  HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet, originalWorkbook, list, kit, doc);
		this.textSheet = sheet;

		checkMergedCells();
		checkEmptyRows();

		if (printFormatError() == false) {
			checkModifiedHeader();
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

	public void checkColorValid() {

		int rowIndex;

		for (rowIndex = 4; rowIndex <= textSheet.getLastRowNum(); rowIndex++) {

			Row row = textSheet.getRow(rowIndex);
			matchColor(row.getCell(19).toString(),"T", rowIndex);

		}
	}
}
