import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;

import org.apache.poi.ss.usermodel.Sheet;

public class ValidateRasterSymbolizer extends CommonValidator{

	private Sheet rasterSheet;

	public ValidateRasterSymbolizer(Sheet sheet, HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet, kit, doc);
		this.rasterSheet = sheet;
		
		checkMergedCells();
		checkEmptyRows();

		if (printFormatError() == false) {
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
