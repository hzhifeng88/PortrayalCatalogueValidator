import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ValidateRasterSymbolizer extends CommonValidator{

	private Sheet rasterSheet;

	public ValidateRasterSymbolizer(Sheet sheet, Workbook originalWorkbook, HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet,originalWorkbook, kit, doc);
		this.rasterSheet = sheet;
		
		checkMergedCells();
		checkEmptyRows();

		if (printFormatError() == false) {
			checkModifiedHeader();
			checkStyleID();
			checkMissingAttributes();
			printValueError();
		}
	}

	public void checkStyleID() {

		for (int rowIndex = 4; rowIndex <= rasterSheet.getLastRowNum(); rowIndex++) {

			checkIDAndDuplicate('R', "F", rowIndex, 5);
		}
	}
}
