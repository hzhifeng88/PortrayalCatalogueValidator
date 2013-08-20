import java.util.ArrayList;

import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ValidateLineSymbolizer extends CommonValidator{

	private Sheet lineSheet;

	public ValidateLineSymbolizer(Sheet sheet, Workbook originalWorkbook, ArrayList<String> list ,  HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet, originalWorkbook, list, kit, doc);
		this.lineSheet = sheet;

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

		for (int rowIndex = 4; rowIndex <= lineSheet.getLastRowNum(); rowIndex++) {

			checkIDAndDuplicate('L', "F", rowIndex, 5);
		}
	}

	public void checkColorValid() {

		int rowIndex;

		for (rowIndex = 4; rowIndex <= lineSheet.getLastRowNum(); rowIndex++) {

			Row row = lineSheet.getRow(rowIndex);
			matchColor(row.getCell(9).toString(), "J", rowIndex);
			matchColor(row.getCell(25).toString(), "Z", rowIndex);
		}
	}
}
