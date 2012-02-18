package SpreadSheet;

/*
 * MainApp.java
 *
 * Created on 17 November 2008, 14:32
 *
 */

/**
 *
 * @author Peter Souter
 */

//Import tools
import java.awt.*;
import java.awt.event.*;
import javax.swing.*;
import javax.swing.event.*;
import java.util.*;
import java.util.regex.*;

public class MainApp extends JPanel implements MouseListener {
	private JPanel NewPane;
	private JSplitPane MainSplitPane;
	private JTextField TextArea;
	private JLabel CellContentsLabel;
	private JLabel BlankCell;
	private JLabel CellType;
	private ArrayList<SpreadCell> RefreshList;
	private SpreadCell CurrentlySelected;
	private SpreadCell aSpreadSheetGrid[][];
	private int Columns;
	private int Rows;

	// Starts the program
	private static void run() {

		// Main window setup
		JFrame MainFrame = new JFrame("JavaSpreadsheet 0.1");
		MainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		MainApp Coursework = new MainApp();
		MainFrame.getContentPane().add(Coursework.MainSplitPane);

		// Add the menu area
		JMenuBar MainMenuBar = new JMenuBar();
		JMenu FileMenu = new JMenu("File");
		MainMenuBar.add(FileMenu);
		JMenuItem MenuItemExit = new JMenuItem("Exit");
		JMenuItem MenuItemNew = new JMenuItem("New");
		FileMenu.add(MenuItemNew);
		FileMenu.add(MenuItemExit);
		MainFrame.setJMenuBar(MainMenuBar);

		// New listener
		ActionListener NewListener = new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				run();
			}
		};

		// Exit listener
		ActionListener ExitListener = new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				System.exit(0);
			}
		};

		// Add actionlistenders
		MenuItemNew.addActionListener(NewListener);
		MenuItemExit.addActionListener(ExitListener);

		// Show the main window
		MainFrame.pack();
		MainFrame.setVisible(true);

	}

	// Main Method
	public static void main(String[] args) {
		run();
	}

	// Sets up main components
	public MainApp() {

		//
		// TEXT PANEL CONSTRUCTION
		//

		// Create Panel for the cell textbox
		JPanel TextFieldPane = new JPanel(new BorderLayout());

		TextFieldPane.setLayout(new GridBagLayout());
		GridBagConstraints XYGridTextArea = new GridBagConstraints();
		JLabel CellLabel = new JLabel("Cell:");
		XYGridTextArea.weightx = 0.2;
		XYGridTextArea.fill = GridBagConstraints.HORIZONTAL;
		XYGridTextArea.gridx = 0;
		XYGridTextArea.gridy = 0;
		TextFieldPane.add(CellLabel, XYGridTextArea);

		TextArea = new JTextField();
		XYGridTextArea.weightx = 2;
		XYGridTextArea.fill = GridBagConstraints.HORIZONTAL;
		XYGridTextArea.gridx = 1;
		XYGridTextArea.gridy = 0;
		TextFieldPane.add(TextArea, XYGridTextArea);

		JButton GoButton = new JButton("Ok");
		XYGridTextArea.weightx = 1;
		XYGridTextArea.gridx = 2;
		XYGridTextArea.gridy = 0;
		TextFieldPane.add(GoButton, XYGridTextArea);

		// Defines what the type of the text is
		// If the text is a formula, defines the formula
		CellType = new JLabel("NONE SELECTED");
		CellType.setHorizontalAlignment(SwingConstants.CENTER);
		XYGridTextArea.weightx = 0.2;
		XYGridTextArea.gridx = 3;
		XYGridTextArea.gridy = 0;
		TextFieldPane.add(CellType, XYGridTextArea);

		// Action Listener for Go Button
		ActionListener GoButtonListener = new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				String s = TextArea.getText();
				SpreadCellTextSet(s);
			}
		};
		GoButton.addActionListener(GoButtonListener);
		TextFieldPane.setBorder(BorderFactory
				.createTitledBorder("Cell Information"));

		//
		// SPREADSHEET PANEL CONSTRUCTION
		//

		// Create Panel for Spreasheet area
		JPanel SpreadSheetGen = new JPanel(new BorderLayout());

		// Set up the GridBag layout for the cells
		SpreadSheetGen.setLayout(new GridBagLayout());
		GridBagConstraints XYGrid = new GridBagConstraints();
		XYGrid.fill = GridBagConstraints.HORIZONTAL;

		// Prompt the user for input on the number of columns and rows
		Columns = Integer.parseInt(JOptionPane.showInputDialog(
				"Number of Columns", "30"));
		Rows = Integer.parseInt(JOptionPane.showInputDialog("Number of Rows",
				"50"));

		// Double array to make finding cells easier
		aSpreadSheetGrid = new SpreadCell[Rows + 1][Columns + 1];

		// Spreadcell arraylist to refresh cell refrences
		RefreshList = new ArrayList<SpreadCell>();

		// Upper right area, blank XY meetup
		BlankCell = new JLabel("      ");
		BlankCell.setBorder(BorderFactory.createRaisedBevelBorder());
		BlankCell.setBackground(Color.lightGray);
		BlankCell.setPreferredSize(new Dimension(60, 20));
		BlankCell.setOpaque(true);
		XYGrid.weightx = 1;
		XYGrid.fill = GridBagConstraints.HORIZONTAL;
		XYGrid.gridx = 0;
		XYGrid.gridy = 0;
		SpreadSheetGen.add(BlankCell, XYGrid);

		// Columns
		// For loop to create the right amount of Columns, as specified by the
		// user
		for (int i = 1; i <= Columns; i++) {
			// Horizontal ie Columns (A..ZZ)
			SpreadCell aSpreadColumn = new SpreadCell();
			aSpreadSheetGrid[0][i] = aSpreadColumn;
			aSpreadColumn.Row = 0;
			aSpreadColumn.Column = i;
			aSpreadColumn.setText(alphabet(i));
			aSpreadColumn.setBorder(BorderFactory.createRaisedBevelBorder());
			aSpreadColumn.setBackground(Color.lightGray);
			aSpreadColumn.setForeground(Color.black);
			aSpreadColumn.setPreferredSize(new Dimension(60, 20));
			aSpreadColumn.setOpaque(true);
			XYGrid.weightx = 1;
			XYGrid.fill = GridBagConstraints.HORIZONTAL;
			XYGrid.gridx = i;
			XYGrid.gridy = 0;
			SpreadSheetGen.add(aSpreadColumn, XYGrid);
		}

		// Rows
		// For loop to create the right amount of rows, as specified by the user
		for (int i = 1; i <= Rows; i++) {
			// Vertical ie Rows (0..999)
			SpreadCell aSpreadRow = new SpreadCell();
			aSpreadSheetGrid[i][0] = aSpreadRow;
			aSpreadRow.Row = i;
			aSpreadRow.Column = 0;
			aSpreadRow.setText("" + i);
			aSpreadRow.setBorder(BorderFactory.createRaisedBevelBorder());
			aSpreadRow.setBackground(Color.lightGray);
			aSpreadRow.setForeground(Color.black);
			aSpreadRow.setPreferredSize(new Dimension(60, 20));
			aSpreadRow.setOpaque(true);
			XYGrid.weightx = 1;
			XYGrid.fill = GridBagConstraints.HORIZONTAL;
			XYGrid.gridx = 0;
			XYGrid.gridy = i;
			SpreadSheetGen.add(aSpreadRow, XYGrid);
		}

		// For loop to create the right amount of cells
		for (int j = 1; j <= Columns; j++) {
			for (int i = 1; i <= Rows; i++) {
				// Fill cells
				SpreadCell aSpreadCell = new SpreadCell();
				aSpreadSheetGrid[i][j] = aSpreadCell;
				aSpreadCell.Row = i;
				aSpreadCell.Column = j;
				aSpreadCell.setBorder(BorderFactory.createRaisedBevelBorder());
				aSpreadCell.setBackground(Color.white);
				aSpreadCell.setForeground(Color.black);
				aSpreadCell.setPreferredSize(new Dimension(60, 20));
				aSpreadCell.setOpaque(true);
				aSpreadCell.addMouseListener(this);
				XYGrid.weightx = 1;
				XYGrid.fill = GridBagConstraints.HORIZONTAL;
				XYGrid.gridx = j;
				XYGrid.gridy = i;
				SpreadSheetGen.add(aSpreadCell, XYGrid);
			}
		}

		SpreadSheetGen.setBorder(BorderFactory
				.createTitledBorder("Spreadsheet"));

		// Add scrollbars to both panels
		JScrollPane TextFieldPaneScrollPane = new JScrollPane(TextFieldPane);
		JScrollPane SpreadSheetScrollPane = new JScrollPane(SpreadSheetGen);

		// Add the two panels to the SplitPane
		MainSplitPane = new JSplitPane(JSplitPane.VERTICAL_SPLIT,
				TextFieldPaneScrollPane, SpreadSheetScrollPane);
		MainSplitPane.setDividerLocation(60);

		// Set prefered sizes
		MainSplitPane.setDividerSize(3);
		MainSplitPane.setPreferredSize(new Dimension(500, 300));

	}

	// Regular Expression tester
	public String regextester(String s) {
		if (s.equals(""))
			return "<EMPTY>";
		if (s.matches("[0-9]+"))
			return "INT";
		else if (s.matches("[0-9]+[.][0-9]+"))
			return "FLOAT";
		else if (s.matches("[=][A-Z]+[0-9]+"))
			return "FORMULAE";
		else
			return "STRING";
	}

	public void AlignmentTester(String s, SpreadCell x) {
		if (s.matches("[0-9]+"))
			x.setHorizontalAlignment(SwingConstants.RIGHT);
		else if (s.matches("[0-9]+[.][0-9]+"))
			x.setHorizontalAlignment(SwingConstants.RIGHT);
		else
			x.setHorizontalAlignment(SwingConstants.LEFT);
	}

	// Converts numbers to column headers (ie. 1 = A...26 = Z, 27 = AA)
	public String alphabet(long i) {
		// A recursive statement for turning numbers into letters
		// Used for making column headers
		if (i == 0)
			return "Z";
		if (i == 1)
			return "A";
		if (i == 2)
			return "B";
		if (i == 3)
			return "C";
		if (i == 4)
			return "D";
		if (i == 5)
			return "E";
		if (i == 5)
			return "E";
		if (i == 6)
			return "F";
		if (i == 7)
			return "G";
		if (i == 8)
			return "H";
		if (i == 9)
			return "I";
		if (i == 10)
			return "J";
		if (i == 11)
			return "K";
		if (i == 12)
			return "L";
		if (i == 13)
			return "M";
		if (i == 14)
			return "N";
		if (i == 15)
			return "O";
		if (i == 16)
			return "P";
		if (i == 17)
			return "Q";
		if (i == 18)
			return "R";
		if (i == 19)
			return "S";
		if (i == 20)
			return "T";
		if (i == 21)
			return "U";
		if (i == 22)
			return "V";
		if (i == 23)
			return "W";
		if (i == 24)
			return "X";
		if (i == 25)
			return "Y";
		if (i == 26)
			return "Z";
		if (i > 26) {
			long ColFirst;
			ColFirst = i / 26;
			if ((i % 26) == 0) {
				ColFirst -= 1;
			}
			return alphabet(ColFirst) + alphabet(i % 26);
		}
		return "?";
	}

	// Converts letters to numbers (ie. A = 1...Z = 26)
	public int alphatonum(String a) {
		// A recursive statement for turning numbers into letters
		// Used for making column headers
		if (a.equals("A"))
			return 1;
		if (a.equals("B"))
			return 2;
		if (a.equals("C"))
			return 3;
		if (a.equals("D"))
			return 4;
		if (a.equals("E"))
			return 5;
		if (a.equals("F"))
			return 6;
		if (a.equals("G"))
			return 7;
		if (a.equals("H"))
			return 8;
		if (a.equals("I"))
			return 9;
		if (a.equals("J"))
			return 10;
		if (a.equals("K"))
			return 11;
		if (a.equals("L"))
			return 12;
		if (a.equals("M"))
			return 13;
		if (a.equals("N"))
			return 14;
		if (a.equals("O"))
			return 15;
		if (a.equals("P"))
			return 16;
		if (a.equals("Q"))
			return 17;
		if (a.equals("R"))
			return 18;
		if (a.equals("S"))
			return 19;
		if (a.equals("T"))
			return 20;
		if (a.equals("U"))
			return 21;
		if (a.equals("V"))
			return 22;
		if (a.equals("W"))
			return 23;
		if (a.equals("X"))
			return 24;
		if (a.equals("Y"))
			return 25;
		if (a.equals("Z"))
			return 26;
		if (a.length() >= 2) {
			String firstone = "" + a.charAt(0);
			String secoundone = "" + a.charAt(1);
			return 26 * alphatonum(firstone) + alphatonum(secoundone);
		}
		return 0;
	}

	// Checks if the input is formulae
	public String isItFormulae(String y) {
		if (y.matches("[=][A-Z]+[0-9]+"))
			return SetFormula(y);
		else
			return y;
	}

	public int getColumn(SpreadCell e) {
		return e.Column;

	}

	public int getRow(SpreadCell e) {
		return e.Row;
	}

	public void SpreadCellTextSet(String s) {
		RefreshReference();
		if (CurrentlySelected == null) {
			CellType.setText("ERROR: Select Cell first");
		}
		// If the textfield is empty
		else if (s.equals("")) {
			CellType.setText("<EMPTY>");
			CurrentlySelected.setText(s);
			CurrentlySelected.Formulae = null;
			CurrentlySelected.isFormulae = false;
		}
		// If the textfield is integer
		else if (s.matches("[0-9]+")) {
			CurrentlySelected.setText(s);
			CurrentlySelected.setHorizontalAlignment(SwingConstants.RIGHT);
			CellType.setText("INT");
			CurrentlySelected.CellType = "INT";
			CurrentlySelected.Formulae = null;
			CurrentlySelected.isFormulae = false;
		}
		// If the text entered is a float
		else if (s.matches("[0-9]+[.][0-9]+")) {
			CurrentlySelected.setHorizontalAlignment(SwingConstants.RIGHT);
			CurrentlySelected.setText(s);
			CellType.setText("FLOAT");
			CurrentlySelected.CellType = "FLOAT";
			CurrentlySelected.Formulae = null;
			CurrentlySelected.isFormulae = false;
		}
		// If the text entered is an Equals formulae
		else if (s.matches("[=][A-Z]+[0-9]+")) {
			CurrentlySelected.Formulae = s;
			CurrentlySelected.isFormulae = true;
			CurrentlySelected.setText("");
			String FormText = FormulaeToCell(s).getText();
			CurrentlySelected.setText(FormText);
			AlignmentTester(FormText, CurrentlySelected);
			TextArea.setText(s);
			RefreshList.add(CurrentlySelected);
			CurrentlySelected.CellType = "FORMULAE";
			CellType.setText("FORMULAE: " + s);
			if (FormulaeToCell(s) == CurrentlySelected) {
				CurrentlySelected.setText("<SelfReferingCell>");
			}
		} else {
			// If the text entered is a string
			CurrentlySelected.setText(s);
			CurrentlySelected.setHorizontalAlignment(SwingConstants.LEFT);
			CellType.setText("STRING");
			CurrentlySelected.CellType = "STRING";
			CurrentlySelected.Formulae = null;
			CurrentlySelected.isFormulae = false;
		}
		RefreshReference();
	}

	public String SetFormula(String s) {
		SpreadCell GetCell = FormulaeToCell(s);
		String NewText = GetCell.getText();
		return NewText;
	}

	// Refreshes cells contents to reflect formulae changes
	public void RefreshReference() {
		// Extra for loop enables formulas refering to formulas to refresh at
		// the same time
		for (int j = 0; j < RefreshList.size(); j++) {
			for (int i = 0; i < RefreshList.size(); i++) {
				if (RefreshList.get(i) != null
						&& RefreshList.get(i).Formulae != null) {
					String Formulae = RefreshList.get(i).Formulae;
					SpreadCell GetCell = FormulaeToCell(Formulae);
					RefreshList.get(i).setText(GetCell.getText());
					AlignmentTester(GetCell.getText(), RefreshList.get(i));
				}
			}
		}
	}

	// Finds the cell a Formulae is refrencing
	public SpreadCell FormulaeToCell(String s) {
		String InputString = s;
		InputString = InputString.replaceAll("=", "");
		String RegEx = "([A-Z]+)([0-9]+)";
		Pattern FormulaePattern = Pattern.compile(RegEx);
		Matcher FormulaeMatch = FormulaePattern.matcher(InputString);
		if (FormulaeMatch.matches()) {
			String aColumn = FormulaeMatch.group(1);
			String aRow = FormulaeMatch.group(2);
			return aSpreadSheetGrid[Integer.parseInt(aRow)][alphatonum(aColumn)];
		}
		// It should not reach this stage, but a return statement must be given
		return aSpreadSheetGrid[1][1];
	}

	// Mouse listeners
	public void mousePressed(MouseEvent e) {
		RefreshReference();
		String RegexTest = "";
		SpreadCell MouseClickedCell = (SpreadCell) e.getSource();
		int SelectedRow = getRow(MouseClickedCell);
		int SelectedColumn = getColumn(MouseClickedCell);
		// De-selecting the cell by clicking on a cell already highlighted
		if (CurrentlySelected == MouseClickedCell) {
			CellType.setText("NONE SELECTED");
			aSpreadSheetGrid[SelectedRow][0].setBackground(Color.lightGray);
			aSpreadSheetGrid[0][SelectedColumn].setBackground(Color.lightGray);
			CurrentlySelected
					.setBorder(BorderFactory.createRaisedBevelBorder());
			CurrentlySelected.setBackground(Color.WHITE);
			MouseClickedCell.setBorder(BorderFactory.createRaisedBevelBorder());
			CurrentlySelected = null;
		}
		// Selecting a cell when no other cell has been highlighted
		else if (CurrentlySelected == null) {
			RegexTest = regextester(MouseClickedCell.getText());
			CellType.setText(RegexTest);
			aSpreadSheetGrid[SelectedRow][0].setBackground(Color.YELLOW);
			aSpreadSheetGrid[0][SelectedColumn].setBackground(Color.YELLOW);
			MouseClickedCell.setBackground(Color.YELLOW);
			MouseClickedCell.setBorder(BorderFactory.createLineBorder(
					Color.BLACK, 2));
			CurrentlySelected = MouseClickedCell;
			if (MouseClickedCell.isFormulae == true) {
				TextArea.setText(MouseClickedCell.Formulae);
				CellType.setText("FORMULAE: " + MouseClickedCell.Formulae);
			}
			if (MouseClickedCell.isFormulae != true) {
				TextArea.setText("");
				CellType.setText(regextester(MouseClickedCell.getText()));
			}
		}
		// Selecting a cell when another cell has been highlighted
		else if (CurrentlySelected != null) {
			int CurrentRow = CurrentlySelected.Row;
			int CurrentColumn = CurrentlySelected.Column;
			RegexTest = regextester(MouseClickedCell.getText());
			CellType.setText(RegexTest);
			aSpreadSheetGrid[CurrentRow][0].setBackground(Color.lightGray);
			aSpreadSheetGrid[0][CurrentColumn].setBackground(Color.lightGray);
			aSpreadSheetGrid[SelectedRow][0].setBackground(Color.YELLOW);
			aSpreadSheetGrid[0][SelectedColumn].setBackground(Color.YELLOW);
			CurrentlySelected
					.setBorder(BorderFactory.createRaisedBevelBorder());
			CurrentlySelected.setBackground(Color.WHITE);
			MouseClickedCell.setBackground(Color.YELLOW);
			MouseClickedCell.setBorder(BorderFactory.createLineBorder(
					Color.BLACK, 2));
			CurrentlySelected = MouseClickedCell;
			if (MouseClickedCell.isFormulae == true) {
				TextArea.setText(MouseClickedCell.Formulae);
				CellType.setText("FORMULAE: " + MouseClickedCell.Formulae);
			}
			if (MouseClickedCell.isFormulae != true) {
				TextArea.setText("");
				CellType.setText(regextester(MouseClickedCell.getText()));
			}
		}
	}

	// Although not being used, the rest have to be declared as the program uses
	// the interface
	public void mouseReleased(MouseEvent e) {
	}

	public void mouseEntered(MouseEvent e) {
	}

	public void mouseExited(MouseEvent e) {
	}

	public void mouseClicked(MouseEvent e) {
	}

}