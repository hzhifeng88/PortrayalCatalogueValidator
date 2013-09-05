import java.io.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.awt.*;
import java.awt.Color;
import java.awt.event.*;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.text.BadLocationException;
import javax.swing.text.html.*;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

public class mainPCV extends JFrame {

	private static mainPCV mainWindow;
	private boolean countOne = false;
	private boolean hasValidated = false;
	private JPanel northPanel;
	private JFileChooser chooser;
	private String excelFilePath;
	private JButton openFileButton;
	private JButton validateButton;
	private JButton exportCSSButton;
	private JTextField pathTextField;
	private JTextPane errorPane;
	private JScrollPane scrollPane;
	private HTMLEditorKit kit;
	private HTMLDocument doc;
	
	private ValidatePointSymbolizer pointSymbolizer;
	private ValidateLineSymbolizer lineSymbolizer;
	private ValidatePolygonSymbolizer polygonSymbolizer;
	private ValidateTextSymbolizer textSymbolizer;
	private ValidateRasterSymbolizer rasterSymbolizer;
	private ValidateColors colorSymbolizer;
	
	// ApachePOI (reading of excel)
	private Workbook workbook;
	private Workbook originalWorkbook;
	private ArrayList<String> storeColorID = new ArrayList<String>();
	private ArrayList<String> storeExtraSheets= new ArrayList<String>();

	public mainPCV() {

		createNorthPanel();
		createSouthPanel();
		
		getContentPane().add(northPanel, BorderLayout.NORTH);
		getContentPane().add(scrollPane, BorderLayout.CENTER);
	}

	public static void main(String[] args) {

		try {
			UIManager.setLookAndFeel("com.jtattoo.plaf.texture.TextureLookAndFeel");
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}

		mainWindow = new mainPCV();
		mainWindow.setTitle("Portrayal Catalogue Validator");
		mainWindow.setSize(530, 560);
		mainWindow.setResizable(false);
		mainWindow.setVisible(true);

		// Set to center of the screen
		Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
		int framePosX = (screenSize.width - mainWindow.getWidth()) / 2;
		int framePosY = (screenSize.height - mainWindow.getHeight()) / 2;
		mainWindow.setLocation(framePosX, framePosY);

		mainWindow.getContentPane();
		mainWindow.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	}

	public void createNorthPanel() {

		northPanel = new JPanel();
		northPanel.setPreferredSize(new Dimension(600, 100));
		northPanel.setBorder(BorderFactory.createTitledBorder("<html><font size = 4> <font color=#0B612D>Select an Excel File (only .xlsx extension)</font color></font></html>"));

		pathTextField = new JTextField();
		pathTextField.setEditable(false);
		pathTextField.setPreferredSize(new Dimension(400, 30));

		openFileButton = new JButton(" ... ");
		openFileButton.addActionListener(new ButtonHandler());

		validateButton = new JButton(" Validate ");
		validateButton.addActionListener(new ButtonHandler());
		
		exportCSSButton = new JButton(" Export to CartoCSS ");
		exportCSSButton.addActionListener(new ButtonHandler());

		northPanel.add(pathTextField);
		northPanel.add(openFileButton);
		northPanel.add(validateButton);
		northPanel.add(exportCSSButton);

		chooser = new JFileChooser();
		chooser.setDialogTitle("Select an Excel File (only .xlsx extension)");
		chooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx"));
	}
	
	public void createSouthPanel(){
		
		errorPane = new JTextPane();
		errorPane.setOpaque(false);
		kit = new HTMLEditorKit();
		doc = new HTMLDocument();
		errorPane.setEditorKit(kit);
		errorPane.setDocument(doc);
		errorPane.setEditable(false);
		errorPane.setSize(450, 450);
		errorPane.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
		
		JViewport viewport = new JViewport() {
			public void paintComponent(Graphics og) {
				super.paintComponent(og);
				Graphics2D g = (Graphics2D) og;
				g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
				GradientPaint gradient = new GradientPaint(0, 0, new Color(247, 237, 204), 0, getHeight(), Color.WHITE, true);
				g.setPaint(gradient);
				g.fillRoundRect(0, 0, getWidth(), getHeight(), 50, 50);
			}
		};
		viewport.add(errorPane);
		scrollPane = new JScrollPane();
		scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
		scrollPane.setViewport(viewport);
	}

	public void enableWindows() {
		mainWindow.setEnabled(true);    
	} 
	   
	public void initializeRead() {

		workbook = null;

		try {
			workbook = WorkbookFactory.create(new FileInputStream(excelFilePath));
			
			if(countOne == false){
				originalWorkbook = workbook;
				countOne = true;
			}
			
			checkExtraSheets();
			readColorsSheet();
			
			pointSymbolizer = new ValidatePointSymbolizer(workbook.getSheetAt(0), originalWorkbook, storeColorID, kit, doc);
			lineSymbolizer = new ValidateLineSymbolizer(workbook.getSheetAt(1), originalWorkbook, storeColorID, kit, doc);
			polygonSymbolizer = new ValidatePolygonSymbolizer(workbook.getSheetAt(2), originalWorkbook, storeColorID, kit, doc);
			textSymbolizer = new ValidateTextSymbolizer(workbook.getSheetAt(3), originalWorkbook, storeColorID, kit, doc);
			rasterSymbolizer = new ValidateRasterSymbolizer(workbook.getSheetAt(4), originalWorkbook, kit, doc);
			colorSymbolizer = new ValidateColors(workbook.getSheetAt(5), originalWorkbook, kit, doc);
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public void checkExtraSheets() {
		
		String tempSheetName;
		storeExtraSheets.clear();
		
		for(int countSheet = 0; countSheet < workbook.getNumberOfSheets(); countSheet++){
			
			Sheet tempSheet = workbook.getSheetAt(countSheet);
			tempSheetName = tempSheet.getSheetName();
			
			  switch(tempSheetName) {
	            case "PointSymbolizer": 
	            	break;
	            case "LineSymbolizer":  
	            	break;
	            case "PolygonSymbolizer":  
	            	break;
	            case "TextSymbolizer": 
	            	break;
	            case "RasterSymbolizer": 
	            	break;
	            case "Colors": 
	            	break;	   
	            default: storeExtraSheets.add(tempSheetName);
	                break;
	        }
		}
		
		try {
			if(storeExtraSheets.isEmpty() == false){
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Extra sheets found: <font color=#088542>" + storeExtraSheets + "<br><br></font color></font>", 0, 0, null);
			}
		} catch (BadLocationException e1) {
			e1.printStackTrace();
		} catch (IOException e1) {
			e1.printStackTrace();
		}
	}

	public void readColorsSheet() {

		Sheet sheet = workbook.getSheetAt(5);

		for (int rowIndex = 4; rowIndex <= sheet.getLastRowNum(); rowIndex++) {

			Row row = sheet.getRow(rowIndex);
			storeColorID.add(row.getCell(0).toString());
		}
	}

	public void exportToCartoCSS() {
		
		// Display export status
		mainWindow.setEnabled(false);   
		
		ExportReport cartoReport = new ExportReport(mainWindow);
		
		cartoReport.writeHeader("PointSymbolizer");
		if(pointSymbolizer.getHasError() == true || colorSymbolizer.getHasError() == true){
			cartoReport.writeTextToReport("Unable to export! Sheet contains error(s).");
		}else {
			
		}
		
		cartoReport.writeHeader("LineSymbolizer");
		if(lineSymbolizer.getHasError() == true || colorSymbolizer.getHasError() == true){
			cartoReport.writeTextToReport("Unable to export! Sheet contains error(s).");
		}else {
			new LineToCartoCSS(workbook.getSheetAt(1), cartoReport, workbook.getSheetAt(5));
		}
		
		cartoReport.writeHeader("PolygonSymbolizer");
		if(polygonSymbolizer.getHasError() == true || colorSymbolizer.getHasError() == true){
			cartoReport.writeTextToReport("Unable to export! Sheet contains error(s).");
		}else {
			
		}
		
		cartoReport.writeHeader("TextSymbolizer");
		if(textSymbolizer.getHasError() == true || colorSymbolizer.getHasError() == true){
			cartoReport.writeTextToReport("Unable to export! Sheet contains error(s).");
		}else {
			new TextToCartoCSS(workbook.getSheetAt(3), cartoReport, workbook.getSheetAt(5));
		}
		
		cartoReport.writeHeader("RasterSymbolizer");
		if(rasterSymbolizer.getHasError() == true || colorSymbolizer.getHasError() == true){
			cartoReport.writeTextToReport("Unable to export! Sheet contains error(s).");
		}else {
			new RasterToCartoCSS(workbook.getSheetAt(4), cartoReport);
		}
	}
	
	private class ButtonHandler implements ActionListener {

		public void actionPerformed(ActionEvent e) {

			if (e.getSource() == openFileButton) {

				int feedBack = chooser.showOpenDialog(null);

				if (feedBack == JFileChooser.OPEN_DIALOG) {
					excelFilePath = chooser.getSelectedFile().toString();
					pathTextField.setText(excelFilePath);

				}
			} else if (e.getSource() == validateButton) {
				hasValidated =  true;
				errorPane.setText("");

				if (excelFilePath == null) {
					JOptionPane.showMessageDialog(null,"Please select an excel file first!");
				} else {
					String tempString = excelFilePath.substring(excelFilePath.length() - 5, excelFilePath.length());

					if (tempString.equalsIgnoreCase(".xlsx")) {
									
						DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
						Date date = new Date();

						try {
							kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Last validated: <font color=#088542>" + dateFormat.format(date) + "<br></font color></font>", 0, 0, null);
						} catch (BadLocationException e1) {
							e1.printStackTrace();
						} catch (IOException e1) {
							e1.printStackTrace();
						}
							
						initializeRead();
					} else {
						JOptionPane.showMessageDialog(null, "Could not process selected file. Did you select the right file?");
					}
				}
			} else if(e.getSource() == exportCSSButton) {
				
				if(hasValidated == true){
					exportToCartoCSS();
				}
				
			}
		}
	}
}