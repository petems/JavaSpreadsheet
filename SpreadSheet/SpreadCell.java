package SpreadSheet;

import javax.swing.JLabel;

//Create a new class "SpreadCell" that extends JLabel
//This allows easier referencing to cell positions
//Also contains the formula of the cell
public class SpreadCell extends JLabel {

	public int Row;
	public int Column;
	boolean isFormulae;
	public String Formulae;
	public String CellType;

	public SpreadCell() {
		int Row = 0;
		int Column = 0;
		boolean isFormulae = false;
		String Formulae = null;
		String CellType = "EMPTY";
	}

}
