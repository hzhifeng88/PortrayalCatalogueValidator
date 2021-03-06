import java.awt.*;
import java.awt.event.*;
import java.io.IOException;

import javax.swing.*;
import javax.swing.text.BadLocationException;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;

public class ExportReport extends JFrame {
	
    public mainPCV mainWindow; 
    private int framePosX;     
    private int framePosY;   
	private JTextPane reportPane;
	private JScrollPane scrollPane;
	private HTMLEditorKit kit;
	private HTMLDocument doc;
    
	public ExportReport(final mainPCV mainWindow) {
		
		this.mainWindow = mainWindow;       
		setTitle("CartoCSS Report");     
		setSize(450, 450);      
		setVisible(true);       
		setResizable(false); 
		
		//Set to center of the screen  
		Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();  
		framePosX = (screenSize.width - getWidth()) / 2;    
		framePosY = (screenSize.height - getHeight()) / 2;  
		setLocation(framePosX, framePosY);        
		
		createPane();
		getContentPane().add(scrollPane, BorderLayout.CENTER);
		
		addWindowListener(new WindowAdapter() {            
			@Override        
			public void windowClosing(WindowEvent e) {
				mainWindow.enableWindows();         
			}
		}); 
	}
	
	public void createPane() {
		
		reportPane = new JTextPane();
		reportPane.setOpaque(false);
		kit = new HTMLEditorKit();
		doc = new HTMLDocument();
		reportPane.setEditorKit(kit);
		reportPane.setDocument(doc);
		reportPane.setEditable(false);
		reportPane.setSize(450, 450);
		reportPane.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
		
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
		viewport.add(reportPane);
		scrollPane = new JScrollPane();
		scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
		scrollPane.setViewport(viewport);
	}
	
	public void writeHeader(String symbolizer) {
		
		try {
			kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#ED0E3F><br><b><font size = 3>" + symbolizer +" </b></font color></font>", 0, 0,null);
			kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#088542>---------------------------- </font color></font>", 0, 0, null);
		}catch (BadLocationException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public void writeTextToReport(String text) {
		
		try {
			kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3>" + text + "</font color></font>", 0, 0,null);
		}catch (BadLocationException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
