import java.io.*;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class LineToCartoCSS extends CommonExport{

	private Sheet lineSheet;
	private BufferedWriter writer;
	private String currentModel;
	private String currentTopic;
	private String currentClass;
	private String storeCartoCSS = "";
	//private ExportReport cartoReport;

	public LineToCartoCSS(Sheet lineSheet, ExportReport cartoReport, Sheet colorSheet) {

		super(colorSheet);
		this.lineSheet = lineSheet;
		//this.cartoReport = cartoReport;
		exportCSS();
	}
	
	public void exportCSS() {

		Row tempRow = lineSheet.getRow(4); 
		currentModel = tempRow.getCell(0).toString();
		currentTopic = tempRow.getCell(1).toString();

		try {

			writer = new BufferedWriter(new FileWriter(currentModel + " - " + currentTopic + ".mss"));

			for (int rowIndex = 4; rowIndex <= lineSheet.getLastRowNum(); rowIndex++) {

				Row row = lineSheet.getRow(rowIndex);
				String tempModel = row.getCell(0).toString();
				String tempTopic = row.getCell(1).toString();
				currentClass = row.getCell(2).toString();
				String tempGeometryAttr = row.getCell(3).toString();

				if(tempModel.equalsIgnoreCase(currentModel) && tempTopic.equalsIgnoreCase(currentTopic)) {
					storeCartoCSS = "";
					appendLayerConditions(lineSheet.getRow(rowIndex), storeCartoCSS, writer, currentClass);
					appendLayerStyle(lineSheet.getRow(rowIndex), tempGeometryAttr);
					writer.append("\r\n");
				}else {
					writer.close();
					storeCartoCSS = "";
					currentModel = tempModel;
					currentTopic = tempTopic;
					writer = new BufferedWriter(new FileWriter(currentModel + " - " + currentTopic + ".mss"));

					appendLayerConditions(lineSheet.getRow(rowIndex), storeCartoCSS, writer, currentClass);
					appendLayerStyle(lineSheet.getRow(rowIndex), tempGeometryAttr);
					writer.append("\r\n");
				}
			}
			writer.close();
		} catch (IOException e) {
			System.out.print(e.getMessage());
		}
	}
	
	public void appendLayerStyle(Row row, String geometryAttr) throws IOException {

		// Graphic or marker based repetition
		if(row.getCell(18) != null && !row.getCell(18).toString().equalsIgnoreCase("")) {

			// Repetition: Initial Gap
			writer.append("\t" + row.getCell(19) + ";");
			writer.append("\r\n");

			// Repetition: Gap
			writer.append("\t" + row.getCell(20) + ";");
			writer.append("\r\n");
			drawPointSymbol(row);
		}else {
			drawGeometryLines(row);
		}
	}
	
	public void drawGeometryLines(Row row) throws IOException {

		// Pencil based color
		if(row.getCell(11) != null && !row.getCell(11).toString().equalsIgnoreCase("")) {

			if(row.getCell(11).toString().charAt(0) == 'C') {
				String foundColor = referenceColor(row.getCell(11).toString());
				writer.append("\tline-color:" + foundColor + ";");
			}else {
				writer.append("\tline-color:" + row.getCell(11) + ";");
			}

			writer.append("\r\n");
		}else {
			writer.append("\tline-color:#000000;");
			writer.append("\r\n");
		}

		// Pencil color opacity
		if(row.getCell(12) != null && !row.getCell(12).toString().equalsIgnoreCase("")) {
			writer.append("\tline-opacity:" + row.getCell(12) + ";");
			writer.append("\r\n");
		}else {
			writer.append("\tline-opacity:1;");
			writer.append("\r\n");
		}

		// Pencil dash array
		writer.append("\tline-dasharray:" + row.getCell(13) + ";");
		writer.append("\r\n");

		// Pencil dash offset
		writer.append("\tline-dash-offset:" + row.getCell(14) + ";");
		writer.append("\r\n");

		// Pencil width
		if(row.getCell(15) != null && !row.getCell(15).toString().equalsIgnoreCase("")) {
			writer.append("\tline-width:" + row.getCell(15) + ";");
			writer.append("\r\n");
		}else {
			writer.append("\tline-width:1;");
			writer.append("\r\n");
		}

		// Pencil line join
		writer.append("\tline-join:" + row.getCell(16) + ";");
		writer.append("\r\n");

		// Pencil line cap
		writer.append("\tline-cap:" + row.getCell(17) + ";");
		writer.append("\r\n");

	}

	public void drawPointSymbol(Row row) throws IOException {

		// NOT DONE NOT DONE NOT DONE
		// Size
		writer.append("\t" + row.getCell(21) + ";");
		writer.append("\r\n");

		// Rotation
		writer.append("\t" + row.getCell(22) + ";");
		writer.append("\r\n");

		// Anchor points & Displacement
		writer.append("\t" + row.getCell(23) + ";");
		writer.append("\r\n");

		if(row.getCell(26) != null && !row.getCell(26).toString().equalsIgnoreCase("")) {

			// Graphic-based filename
			writer.append("\tline-pattern-file" + row.getCell(26) + ";");
			writer.append("\r\n");

		}else {
			// Marker based on well known symbol
			writer.append("\t" + row.getCell(24) + ";");
			writer.append("\r\n");

			// Marker based on glyph
			writer.append("\t" + row.getCell(25) + ";");
			writer.append("\r\n");
			fillMarkerArea(row);
		}
	}
	
	public void fillMarkerArea(Row row) throws IOException {

		// Solid color based
		if(row.getCell(27) != null && !row.getCell(27).toString().equalsIgnoreCase("")) {
			if(row.getCell(27).toString().charAt(0) == 'C') {

				String foundColor = referenceColor(row.getCell(27).toString());

				if(!foundColor.equalsIgnoreCase("")) {
					writer.append("\tmarker-fill :" + foundColor + ";");
					writer.append("\r\n");
				}
			}else {
				writer.append("\tmarker-fill:" + row.getCell(27) + ";");
				writer.append("\r\n");
			}
		}else {
			writer.append("\tmarker-fill:#808080;");
			writer.append("\r\n");
		}

		// Solid color opacity
		if(row.getCell(28) != null && !row.getCell(28).toString().equalsIgnoreCase("")) {
			writer.append("\tmarker-fill-opacity:" + row.getCell(28) + ";");
			writer.append("\r\n");
		}else {
			writer.append("\tmarker-fill-opacity:1;");
			writer.append("\r\n");
		}

		// Implement Graphic or marker pattern-based area repetition HERE

	}
}
