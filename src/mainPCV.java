import java.io.*;
import java.util.*;
import java.awt.*;
import java.awt.event.*;

import javax.swing.*;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.Workbook;

public class mainPCV extends JFrame {

	// GUI
	private static mainPCV mainWindow;
	private JPanel northPanel;
	private JFileChooser chooser;
	private String excelFilePath;
	private JButton openFileButton;
	private JButton validateButton;
	private JTextField pathTextField;
	private JTextPane errorPane;
	private JScrollPane scrollPane;
	private HTMLEditorKit kit;
	private HTMLDocument doc;

	// ApachePOI (reading of excel)
	private Workbook workbook;
	private ColorsObject ColorsObject;
	private LinkedList<ColorsObject> storeColorsObjectList = new LinkedList<ColorsObject>();

	public mainPCV() {

		createNorthPanel();

		errorPane = new JTextPane();
		kit = new HTMLEditorKit();
		doc = new HTMLDocument();
		errorPane.setEditorKit(kit);
		errorPane.setDocument(doc);
		errorPane.setEditable(false);
		errorPane.setSize(450, 450);
		errorPane.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

		scrollPane = new JScrollPane(errorPane);
		scrollPane
				.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);

		getContentPane().add(northPanel, BorderLayout.NORTH);
		getContentPane().add(scrollPane, BorderLayout.CENTER);
	}

	public static void main(String[] args) {

		try {
			UIManager
					.setLookAndFeel("com.jtattoo.plaf.texture.TextureLookAndFeel");

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
		northPanel.setPreferredSize(new Dimension(600, 110));
		northPanel
				.setBorder(BorderFactory
						.createTitledBorder("<html><font size = 4> <font color=#088542>Select an Excel File</font color></font></html>"));

		pathTextField = new JTextField();
		pathTextField.setEditable(false);
		pathTextField.setPreferredSize(new Dimension(400, 30));

		openFileButton = new JButton(" Browse ");
		openFileButton.addActionListener(new ButtonHandler());

		validateButton = new JButton(" Validate ");
		validateButton.addActionListener(new ButtonHandler());

		northPanel
				.add(new JLabel(
						"<html><font size = 4> <font color=#0E11F5>File Path:</font color></font></html>"));
		northPanel.add(pathTextField);
		northPanel.add(openFileButton);
		northPanel.add(validateButton);

		chooser = new JFileChooser();
	}

	public void initializeRead() {

		workbook = null;

		try {

			workbook = WorkbookFactory
					.create(new FileInputStream(excelFilePath));

			readColorsSheet();

			new ValidatePointSymbolizer(workbook.getSheetAt(0),
					storeColorsObjectList, kit, doc);

			new ValidateLineSymbolizer(workbook.getSheetAt(1),
					storeColorsObjectList, kit, doc);

			new ValidatePolygonSymbolizer(workbook.getSheetAt(2),
					storeColorsObjectList, kit, doc);

			new ValidateTextSymbolizer(workbook.getSheetAt(3),
					storeColorsObjectList, kit, doc);

			new ValidateRasterSymbolizer(workbook.getSheetAt(4), kit, doc);

			new ValidateColors(workbook.getSheetAt(5), kit, doc);

		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	// This function reads in sheet and stores the data(rows) in a object data
	// structure. 1 Row = 1 Object
	public void readColorsSheet() {

		// Get the Colors sheet.
		Sheet sheet = workbook.getSheetAt(5);

		for (int rowIndex = 4; rowIndex <= sheet.getLastRowNum(); rowIndex++) {

			Row row = sheet.getRow(rowIndex);

			// NOTE: use equalsIgnoreCase("") to check for blank cells if needed
			ColorsObject = new ColorsObject(row.getCell(0).toString(), row
					.getCell(1).toString(), row.getCell(2).toString(), row
					.getCell(3).toString());

			storeColorsObjectList.add(ColorsObject);

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

				errorPane.setText("");

				if (excelFilePath == null) {
					JOptionPane.showMessageDialog(null,
							"Please select an excel file first!");
				} else {
					String tempString = excelFilePath.substring(
							excelFilePath.length() - 5, excelFilePath.length());

					if (tempString.equalsIgnoreCase(".xlsx")) {
						initializeRead();
					} else {
						JOptionPane
								.showMessageDialog(null,
										"Could not process selected file. Did you select the right file?");
					}
				}
			}
		}
	}
}