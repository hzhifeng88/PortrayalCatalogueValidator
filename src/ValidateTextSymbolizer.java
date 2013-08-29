import java.util.ArrayList;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;
import org.apache.poi.ss.usermodel.*;

public class ValidateTextSymbolizer extends CommonValidator {

	private Sheet textSheet;

	public ValidateTextSymbolizer(Sheet sheet, Workbook originalWorkbook, ArrayList<String> list,  HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet, originalWorkbook, list, kit, doc);
		this.textSheet = sheet;

		checkMergedCells();
		checkEmptyRows();

		if (printFormatError() == false) {
			checkModifiedHeader();
			performChecks();
			printValueError();
		}
	}
	
	public void performChecks() {

		for (int rowIndex = 4; rowIndex <= textSheet.getLastRowNum(); rowIndex++) {

			Row row = textSheet.getRow(rowIndex);
			
			// Check valid ID and duplicate
			checkIDAndDuplicate('T', "F", rowIndex, 5);
			
			// Check color valid
			matchColor(row.getCell(21).toString(), "V", rowIndex);
			
			// Check missing attributes
			checkMissingAttributes(row, rowIndex);
		}
	}
}
