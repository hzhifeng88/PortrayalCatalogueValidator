import java.util.ArrayList;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ValidateLineSymbolizer extends CommonValidator{

	private Sheet lineSheet;

	public ValidateLineSymbolizer(Sheet sheet, ArrayList<String> list ,  HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet, list, kit, doc);
		this.lineSheet = sheet;

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

		for (int rowIndex = 4; rowIndex <= lineSheet.getLastRowNum(); rowIndex++) {

			checkIDAndDuplicate('L', "F", rowIndex, 5);
		}
	}

	// Column J and Z store color values or references
	public void checkColorValid() {

		int rowIndex;

		for (rowIndex = 4; rowIndex <= lineSheet.getLastRowNum(); rowIndex++) {

			Row row = lineSheet.getRow(rowIndex);
			matchColor(row.getCell(9).toString(), "J", rowIndex);
			matchColor(row.getCell(25).toString(), "Z", rowIndex);
		}
	}
}
