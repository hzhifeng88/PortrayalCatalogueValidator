import java.util.ArrayList;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;
import org.apache.poi.ss.usermodel.*;

public class ValidatePointSymbolizer extends CommonValidator {

	private Sheet pointSheet;
	private boolean sheetCorrect = false;

	public ValidatePointSymbolizer(Sheet pointSheet, Workbook originalWorkbook, ArrayList<String> list, HTMLEditorKit kit, HTMLDocument doc) {

		super(pointSheet, originalWorkbook, list, kit, doc);
		this.pointSheet = pointSheet;

		checkMergedCells();
		checkEmptyRows();

		if (printFormatError() == false) {
			checkModifiedHeader();
			performChecks();
			printValueError();
			
			if(getHasError() == false){
				sheetCorrect = true;
			}else {
				sheetCorrect = false;
			}
		}
	}
	
	public boolean isSheetCorrect() {
		return sheetCorrect;
	}

	public void performChecks() {

		for (int rowIndex = 4; rowIndex <= pointSheet.getLastRowNum(); rowIndex++) {

			Row row = pointSheet.getRow(rowIndex);

			// Check valid ID and duplicate
			checkIDAndDuplicate('P', "F", rowIndex, 5);

			// Check color valid
			if(row.getCell(17) != null && !row.getCell(17).toString().equalsIgnoreCase("")) {
				matchColor(row.getCell(17).toString(), "R", rowIndex);
			}
			if(row.getCell(27) != null && !row.getCell(27).toString().equalsIgnoreCase("")) {
				matchColor(row.getCell(27).toString(), "AB", rowIndex);
			}

			// Check missing attributes
			checkMissingAttributes(row, rowIndex);
		}
	}
}
