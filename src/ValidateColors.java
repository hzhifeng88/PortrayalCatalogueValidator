import java.io.IOException;
import java.util.ArrayList;

import javax.swing.text.BadLocationException;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ValidateColors extends CommonValidator {

	private Sheet colorSheet;
	private ArrayList<String> storeInvalidColorRGB = new ArrayList<String>();
	private HTMLEditorKit kit;
	private HTMLDocument doc;

	public ValidateColors(Sheet sheet, Workbook originalWorkbook, HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet,originalWorkbook, kit ,doc);
		this.colorSheet = sheet;
		this.kit = kit;
		this.doc = doc;

		checkMergedCells();
		checkEmptyRows();

		if (printFormatError() == false) {
			checkModifiedHeader();
			checkColorID();
			checkRGB();
			printValueError();
			printRGBError();
		}
	}

	public void checkColorID() {

		for (int rowIndex = 4; rowIndex <= colorSheet.getLastRowNum(); rowIndex++) {
			
			checkIDAndDuplicate('C', "A", rowIndex, 0);
		}
	}

	public void checkRGB() {
		
		for (int rowIndex = 4; rowIndex <= colorSheet.getLastRowNum(); rowIndex++) {
			
			Row row = colorSheet.getRow(rowIndex);
			
			if(row.getCell(1) != null){
				String tempString = row.getCell(1).toString();
				checkLineBreak(tempString, "B", rowIndex);

				if (!tempString.equalsIgnoreCase("")) {
					if (isRGB(tempString) == false) {
						storeInvalidColorRGB.add("B" + Integer.toString(rowIndex + 1));
					}
				}
			}
		}
	}
	
	public void printRGBError(){
		
		try {
			
			if (storeInvalidColorRGB.isEmpty() == false) {
				hasError = true;
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> sRGB is invalid (Rule 1: Begins with '#', Rule 2: 6 hexadecimal representation)</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeInvalidColorRGB + "</font color></font>", 0, 0, null);
			}
		} catch (BadLocationException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
