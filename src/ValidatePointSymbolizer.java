import java.util.ArrayList;
import java.io.File;

import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;
import org.apache.poi.ss.usermodel.*;

import javax.xml.parsers.*;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.*;

public class ValidatePointSymbolizer extends CommonValidator {

	private Sheet pointSheet;
	private Document doc;

	public ValidatePointSymbolizer(Sheet pointSheet, Workbook originalWorkbook, ArrayList<String> list, HTMLEditorKit kit, HTMLDocument doc) {

		super(pointSheet, originalWorkbook, list, kit, doc);
		this.pointSheet = pointSheet;

		checkMergedCells();
		checkEmptyRows();

		if (printFormatError() == false) {
			checkModifiedHeader();
			performChecks();
			printValueError();
//			writeXML();
		}
	}

	public void performChecks() {

		for (int rowIndex = 4; rowIndex <= pointSheet.getLastRowNum(); rowIndex++) {

			Row row = pointSheet.getRow(rowIndex);

			// Check valid ID and duplicate
			checkIDAndDuplicate('P', "F", rowIndex, 5);

			// Check color valid
			matchColor(row.getCell(17).toString(), "R", rowIndex);
			matchColor(row.getCell(25).toString(), "Z", rowIndex);

			// Check missing attributes
			checkMissingAttributes(row, rowIndex);
			
//			preparePointXML();
		}
	}

//	public void addElement(Document doc, Element graphic){
//		
//		for (int rowIndex = 4; rowIndex <= pointSheet.getLastRowNum(); rowIndex++) {
//			
//			XSSFRow row = pointSheet.getRow(rowIndex);
//			
//			Element styleID = doc.createElement("StyleID");
//			graphic.appendChild(styleID);
//			
//			// Set attribute to styleID element
//			Attr attr = doc.createAttribute("id");
//			attr.setValue(row.getCell(5).toString());
//			styleID.setAttributeNode(attr);
//			
//			Element color = doc.createElement("Color");
//			color.appendChild(doc.createTextNode(row.getCell(17).toString()));
//			styleID.appendChild(color);
//		}
//	}
//	
//	public void preparePointXML() {
//		
//		try {
//			DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
//			DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
//
//			// Root element
//			Document doc = docBuilder.newDocument();
//			Element rootElement = doc.createElement("PointSymbolizer");
//			doc.appendChild(rootElement);
//
//			// Geometry element
//			Element geometry = doc.createElement("Geometry");
//			rootElement.appendChild(geometry);
//
//			// Graphic elements
//			Element graphic = doc.createElement("Graphic");
//			geometry.appendChild(graphic);
//			
//			addElement(doc, graphic);
//		} catch (ParserConfigurationException pce) {
//			pce.printStackTrace();
//		}
//	}
//	
//	public void writeXML() {
//		
//		// Write the content into XML file
//		TransformerFactory transformerFactory = TransformerFactory.newInstance();
//		Transformer transformer;
//		
//		try {
//			transformer = transformerFactory.newTransformer();
//			DOMSource source = new DOMSource(doc);
//			StreamResult result = new StreamResult(new File("PointSheetXML.xml"));
//			transformer.transform(source, result);
//			System.out.println("XML Generated!");
//		} catch (TransformerConfigurationException e) {
//			e.printStackTrace();
//		} catch (TransformerException e) {
//			e.printStackTrace();
//		}
//	}
}
