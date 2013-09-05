import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;
import org.apache.poi.ss.usermodel.*;

public class ValidateRasterSymbolizer extends CommonValidator{

	private Sheet rasterSheet;
	private boolean sheetCorrect = false;

	public ValidateRasterSymbolizer(Sheet sheet, Workbook originalWorkbook, HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet,originalWorkbook, kit, doc);
		this.rasterSheet = sheet;
		
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

		for (int rowIndex = 4; rowIndex <= rasterSheet.getLastRowNum(); rowIndex++) {

			Row row = rasterSheet.getRow(rowIndex);
			
			// Check valid ID and duplicate
			checkIDAndDuplicate('R', "F", rowIndex, 5);

			// Check missing attributes
			checkMissingAttributes(row, rowIndex);
		}
	}
}
