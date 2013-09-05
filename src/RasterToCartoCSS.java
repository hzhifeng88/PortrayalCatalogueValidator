import java.io.*;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class RasterToCartoCSS extends CommonExport{

	private Sheet rasterSheet;
	private BufferedWriter writer;
	private String currentModel;
	private String currentTopic;
	private String currentClass;
	private String storeCartoCSS = "";
	
	public RasterToCartoCSS(Sheet rasterSheet, ExportReport cartoReport) {
		
		super();
		this.rasterSheet = rasterSheet;
		exportCSS();
	}

	public void exportCSS() {
		
		Row tempRow = rasterSheet.getRow(4); 
		currentModel = tempRow.getCell(0).toString();
		currentTopic = tempRow.getCell(1).toString();
		
		try {

			writer = new BufferedWriter(new FileWriter(currentModel + " - " + currentTopic + ".mss"));

			for (int rowIndex = 4; rowIndex <= rasterSheet.getLastRowNum(); rowIndex++) {
				
				Row row = rasterSheet.getRow(rowIndex);
				String tempModel = row.getCell(0).toString();
				String tempTopic = row.getCell(1).toString();
				currentClass = row.getCell(2).toString();
				String tempGeometryAttr = row.getCell(3).toString();
								
				if(tempModel.equalsIgnoreCase(currentModel) && tempTopic.equalsIgnoreCase(currentTopic)) {
					storeCartoCSS = "";
					appendLayerConditions(row, storeCartoCSS, writer, currentClass);
					appendLayerStyle(rasterSheet.getRow(rowIndex), tempGeometryAttr);
					writer.append("\r\n");
				}else {
					writer.close();
					storeCartoCSS = "";
					currentModel = tempModel;
					currentTopic = tempTopic;
					writer = new BufferedWriter(new FileWriter(currentModel + " - " + currentTopic + ".mss"));
					
					appendLayerConditions(row, storeCartoCSS, writer, currentClass);
					appendLayerStyle(rasterSheet.getRow(rowIndex), tempGeometryAttr);
					writer.append("\r\n");
				}
			}
			writer.close();
		} catch (IOException e) {
			System.out.print(e.getMessage());
		}
	}

	public void appendLayerStyle(Row row, String geometryAttr) throws IOException {
		
		// Opacity
		if(row.getCell(11) != null && !row.getCell(11).toString().equalsIgnoreCase("")) {
			writer.append(" {\r\n");
			writer.append("\traster-opacity:" + row.getCell(11) + ";");
			writer.append("\r\n");
		}else {
			writer.append(" {\r\n");
			writer.append("\traster-opacity:1;");
			writer.append("\r\n");
		}
	}
}
