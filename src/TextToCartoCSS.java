import java.io.*;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class TextToCartoCSS extends CommonExport {

	private Sheet textSheet;
	private BufferedWriter writer;
	private String currentModel;
	private String currentTopic;
	private String currentClass;
	private String storeCartoCSS = "";
	private ExportReport cartoReport;

	public TextToCartoCSS(Sheet textSheet, ExportReport cartoReport, Sheet colorSheet){

		super(colorSheet);
		this.textSheet = textSheet;
		this.cartoReport = cartoReport;
		exportCSS();
	}

	public void exportCSS() {

		Row tempRow = textSheet.getRow(4);
		currentModel = tempRow.getCell(0).toString();
		currentTopic = tempRow.getCell(1).toString();

		try {

			writer = new BufferedWriter(new FileWriter(currentModel + " - " + currentTopic + ".mss"));

			for (int rowIndex = 4; rowIndex <= textSheet.getLastRowNum(); rowIndex++) {

				Row row = textSheet.getRow(rowIndex);
				String tempModel = row.getCell(0).toString();
				String tempTopic = row.getCell(1).toString();
				currentClass = row.getCell(2).toString();
				String tempGeometryAttr = row.getCell(3).toString();

				if(tempModel.equalsIgnoreCase(currentModel) && tempTopic.equalsIgnoreCase(currentTopic)) {
					storeCartoCSS = "";
					appendLayerConditions(row, storeCartoCSS, writer, currentClass);
					appendLayerStyle(textSheet.getRow(rowIndex), tempGeometryAttr);
					writer.append("\r\n");
				}else {
					writer.close();
					storeCartoCSS = "";
					currentModel = tempModel;
					currentTopic = tempTopic;
					writer = new BufferedWriter(new FileWriter(currentModel + " - " + currentTopic + ".mss"));

					appendLayerConditions(row, storeCartoCSS, writer, currentClass);
					appendLayerStyle(textSheet.getRow(rowIndex), tempGeometryAttr);
					writer.append("\r\n");
				}
			}
			writer.close();
		} catch (IOException e) {
			System.out.print(e.getMessage());
		}
	}
	
	public void appendLayerStyle(Row row, String geometryAttr) throws IOException {

		labelingGeneral(row);

		if(geometryAttr.equalsIgnoreCase("point")) {
			labelingToPoint(row);
		} else if(geometryAttr.equalsIgnoreCase("line")){
			labelingToLine(row);
		}

		fillArea(row);
		writer.append("}\r\n");
	}

	public void labelingGeneral(Row row) throws IOException {

		// Label text
		writer.append(" {\r\n");
		writer.append("\ttext-name:\"[" + row.getCell(11) + "]\";");
		writer.append("\r\n");

		// Default font: Times new Roman Regular
		String fontFamily = row.getCell(12).toString();
		String tempFontArray[] = fontFamily.split(",");
		if(referenceFont(tempFontArray[0]) == true) {
			writer.append("\ttext-face-name:\"" + tempFontArray[0] + "\";");
			writer.append("\r\n");

		}else {
			writer.append("\ttext-face-name:\"Times New Roman Regular\";");
			writer.append("\r\n");
			cartoReport.writeTextToReport("Default font used: Times New Roman Regular");
		}

		// Halo color, Radius
		String haloColorRadius = row.getCell(13).toString();
		String tempHaloArray[] = haloColorRadius.split(",");
		writer.append("\ttext-halo-radius:" + tempHaloArray[1] + ";");
		writer.append("\r\n");
		writer.append("\ttext-halo-fill:" + tempHaloArray[0] + ";");
		writer.append("\r\n");

		// Font size
		if(row.getCell(14) != null && !row.getCell(14).toString().equalsIgnoreCase("")) {
			Double tempDouble = new Double(row.getCell(14).toString());
			writer.append("\ttext-size:" + (int)Math.round(tempDouble) + ";");
			writer.append("\r\n");
		}else {
			writer.append("\ttext-size: 10;");
			writer.append("\r\n");
		}
	}

	public void labelingToPoint(Row row) throws IOException {

		// Rotation
		writer.append("\ttext-orientation:" + row.getCell(15) + ";");
		writer.append("\r\n");

		// Anchor point
		String anchorPoint = row.getCell(16).toString();
		String tempAnchorArray[] = anchorPoint.split(",");
		writer.append("\t" + tempAnchorArray[0]  + ";");
		writer.append("\r\n");
		writer.append("\t" + tempAnchorArray[1]  + ";");
		writer.append("\r\n");

		// X,Y displacement
		String xyDisplacement = row.getCell(17).toString();
		String tempDisArray[] = xyDisplacement.split(",");
		writer.append("\ttext-dx:" + tempDisArray[0]  + ";");
		writer.append("\r\n");
		writer.append("\ttext-dy:" + tempDisArray[1]  + ";");
		writer.append("\r\n");

	}

	public void labelingToLine(Row row) throws IOException {

		// Perpendicular offset
		if(row.getCell(18) != null && !row.getCell(18).toString().equalsIgnoreCase("")) {
			writer.append("\t" + row.getCell(18)  + ";");
			writer.append("\r\n");
		}else {
			writer.append("\t0;");
			writer.append("\r\n");
		}

		// Repeated gaps
		if(row.getCell(19) != null && !row.getCell(19).toString().equalsIgnoreCase("")) {
			String repeatedGaps = row.getCell(19).toString();
			String tempGapArray[] = repeatedGaps.split(",");
			writer.append("\t" + tempGapArray[0]  + ";");
			writer.append("\r\n");
			writer.append("\t" + tempGapArray[1]  + ";");
			writer.append("\r\n");
		}

		// Alignment (Geometry or Horizontal)
		if(row.getCell(20) != null && !row.getCell(20).toString().equalsIgnoreCase("")) {
			writer.append("\t" + row.getCell(18)  + ";");
			writer.append("\r\n");
		}else {
			writer.append("\t;");
			writer.append("\r\n");
		}

	}

	public void fillArea(Row row) throws IOException{

		// Solid color based
		if(row.getCell(21) != null && !row.getCell(21).toString().equalsIgnoreCase("")) {
			if(row.getCell(21).toString().charAt(0) == 'C') {

				String foundColor = referenceColor(row.getCell(21).toString());

				if(!foundColor.equalsIgnoreCase("")) {
					writer.append("\ttext-fill:" + foundColor + ";");
					writer.append("\r\n");
				}
			}else {
				writer.append("\ttext-fill:" + row.getCell(21) + ";");
				writer.append("\r\n");
			}
		}else {
			writer.append("\ttext-fill:#808080;");
			writer.append("\r\n");
		}

		// Solid color opacity
		writer.append("\ttext-opacity:" + row.getCell(22) + ";");
		writer.append("\r\n");

	}
}
