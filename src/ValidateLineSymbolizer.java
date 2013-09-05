import java.util.ArrayList;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;
import org.apache.poi.ss.usermodel.*;

public class ValidateLineSymbolizer extends CommonValidator{

	private Sheet lineSheet;
	private boolean sheetCorrect = false;

	public ValidateLineSymbolizer(Sheet sheet, Workbook originalWorkbook, ArrayList<String> list ,  HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet, originalWorkbook, list, kit, doc);
		this.lineSheet = sheet;

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

		for (int rowIndex = 4; rowIndex <= lineSheet.getLastRowNum(); rowIndex++) {

			Row row = lineSheet.getRow(rowIndex);
			
			// Check valid ID and duplicate
			checkIDAndDuplicate('L', "F", rowIndex, 5);
			
			// Check color valid
			if(row.getCell(11) != null && !row.getCell(11).toString().equalsIgnoreCase("")) {
				matchColor(row.getCell(11).toString(), "L", rowIndex);
			}
			if(row.getCell(27) != null && !row.getCell(27).toString().equalsIgnoreCase("")) {
				matchColor(row.getCell(27).toString(), "AB", rowIndex);
			}
			
			// Check missing attributes
			checkMissingAttributes(row, rowIndex);
		}
	}
}
