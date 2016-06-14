package mintec;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

public class MintReader {
	@SuppressWarnings("unused")
	private final int A = 1, B = 2, C = 3, D = 4, E = 5, F = 6, G = 7, H = 8;

	private String name;
	private Date date;
	private ArrayList<String> subjects;
	private double subjectsMean;
	private int subjectsLevel;

	private String projectString = "";
	private int projectLevel;

	private ArrayList<String> activities1;
	private ArrayList<String> activities2;
	private int activityLevel;

	private Sheet sheet;
	private List<Problem> problems;

	public class Problem {
		public final int row;
		public final String column;
		public final String text;
		public final Boolean fatal;

		public Problem(int row, String column, String text, Boolean fatal) {
			super();
			this.row = row;
			this.column = column;
			this.text = text;
			this.fatal = fatal;
		}

		public Problem(int row, int column, String text, Boolean fatal) {
			super();
			String columnLetters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
			this.row = row;
			this.column = "" + columnLetters.charAt(column - 1); // TODO: total implementation
			this.text = text;
			this.fatal = fatal;
		}

	}

	MintReader(FileInputStream file) throws InvalidFormatException, IOException {
		sheet = WorkbookFactory.create(file).getSheetAt(0);
		problems = new ArrayList<>();
		{
			final String expectedVersion = "1.0.0";
			final String readVersion = stringCellAt(A,3);
			if(!expectedVersion.equals(readVersion)) {
				problems.add(new Problem(A, 3, "Inkompatible Formularversion: " + "Version ist " + readVersion + ", erwarte " + expectedVersion, true));
				return;
			}
		}
		{
			final int personColumn = C;
			final int nameRow = 1;
			final int dateRow = 2;

			name = stringCellAt(personColumn, nameRow);
			if (name.equals("")) {
				problems.add(new Problem(nameRow, personColumn, "fehlender Name", true));
			}
			try {
				date = cellAt(personColumn, dateRow).getDateCellValue();
				if (date == null)
					problems.add(new Problem(dateRow, personColumn, "Geburtsdatum fehlt", true));
			} catch (Exception e) {
			    String dateString = stringCellAt(personColumn, dateRow);
				problems.add(new Problem(dateRow, personColumn,
							             "fehlerhafte Datumsangabe: " + dateString, true));
			}
		}

		final int textColumn = B;
		final int gradeColumn = D;
		final int levelColumn = H;

		extractSubjects(textColumn, gradeColumn);
		extractProjects(textColumn, gradeColumn);
		extractActivities(textColumn, levelColumn);
	}

	private void extractActivities(final int textColumn, final int levelColumn) {
		final int activities1startRow = 28;
		final int activities1endRow = activities1startRow + 20;
		final int activities2startRow = 49;
		final int activities2endRow = activities2startRow + 20;
		final int extraRow = 72;

		activityLevel = intCellAt(levelColumn, extraRow);

		activities1 = new ArrayList<>();
		activities2 = new ArrayList<>();

		for(int row = activities1startRow; row < activities1endRow; ++row) {
			String activity = stringCellAt(textColumn, row).trim();
			if(!activity.equals("")) activities1.add(activity);
		}

		for(int row = activities2startRow; row < activities2endRow; ++row) {
			String activity = stringCellAt(textColumn, row).trim();
			if(!activity.equals("")) activities2.add(activity);
		}
	}

	private void extractProjects(final int textColumn, final int gradeColumn) {
		final int subLevelColumn = E;

		final int rowA = 15;
		final int rowB = 17;
		final int rowC = 20;
		final int rowD = 23;

		final int levelA = intCellAt(subLevelColumn, rowA);
		final int levelB = intCellAt(subLevelColumn, rowB);
		final int levelC = intCellAt(subLevelColumn, rowC);
		int levelD = intCellAt(subLevelColumn, rowD);

		if(levelD > 3) {
			problems.add(new Problem(rowD, subLevelColumn, "Stufe des Jugend forscht-Wettbewerbs größer als 3", true));
			return;
		}
		projectLevel = Math.max(levelA, Math.max(levelB, Math.max(levelC, levelD)));

		final int nameRow;

		if (projectLevel == levelA) {
			nameRow = rowA;
			projectString = "Wissenschaftspropädeutisches Fach: ";
		} else if (projectLevel == levelB) {
			nameRow = rowB;
			projectString = "Fachwissenschaftliche Arbeit\n\nFach: ";
		} else if (projectLevel == levelC) {
		    nameRow = rowC;
			projectString = "Besondere Lernleistung\n\nFach: ";
		} else if(projectLevel == levelD) {
			nameRow = rowD;
		} else {
			return;
		}

		if(stringCellAt(textColumn, nameRow).isEmpty()) {
			problems.add(new Problem(nameRow, textColumn, "Fehlende Eingabe", true));
		}
		projectString += stringCellAt(textColumn, nameRow);


		if(nameRow != rowA) {
			if(stringCellAt(textColumn, nameRow + 1).isEmpty()) {
				problems.add(new Problem(nameRow + 1, textColumn, "Fehlende Eingabe", true));
			}
			projectString += "\n\n" + "Thema:\n" + stringCellAt(textColumn, nameRow + 1);
		}

		if(nameRow == rowD) {
			if(stringCellAt(textColumn, nameRow + 2).isEmpty()) {
				problems.add(new Problem(nameRow + 2, textColumn, "Fehlende Eingabe", true));
			}
			projectString += "\n\n" + stringCellAt(textColumn, nameRow + 2);
		} else
			projectString += "\n\n" + "Note: " + Integer.toString(intCellAt(gradeColumn, nameRow));
	}

	private void extractSubjects(final int textColumn, final int gradeColumn) {
		final int meanColumn = E;
		final int meanRow = 7;
		final int twoSubjectsStartRow = 7, threeSubjectsStartRow = 10;

		int startRow = 0, endRow = 0;
		int twoSubjectsFilledFields = 0, threeSubjectsFilledFields = 0;

		for (int column = textColumn; column <= gradeColumn; column += 2) {
			for (int row = twoSubjectsStartRow; row < twoSubjectsStartRow + 2; row += 1) {
				if (!cellEmpty(column, row)) twoSubjectsFilledFields += 1;
			}
			for (int row = threeSubjectsStartRow; row < threeSubjectsStartRow + 3; row += 1) {
				if (!cellEmpty(column, row)) threeSubjectsFilledFields += 1;
			}
		}

		if(twoSubjectsFilledFields == 4) {
			startRow = twoSubjectsStartRow;
			endRow = startRow + 2;
			if(threeSubjectsFilledFields > 0) {
				problems.add(new Problem(threeSubjectsStartRow, textColumn,
						 "Ignoriere überflüssige Eingabe", false));
			}
		} else if(threeSubjectsFilledFields == 6) {
			startRow = threeSubjectsStartRow;
			endRow = startRow + 3;
			if (twoSubjectsFilledFields > 0)
				problems.add(new Problem(twoSubjectsStartRow, textColumn,
										 "Ignoriere überflüssige Eingabe", false));
		} else {
			problems.add(new Problem(A, 4, "Weder zwei Abiturfächer auf erhöhtem Niveau noch ein Abiturfach auf erhöhtem Niveau und zwei weitere Fächer vollständig ausgefüllt.", true));
		}

		subjectsMean = cellAt(meanColumn, meanRow).getNumericCellValue();
		subjects = new ArrayList<>();
		for (int row = startRow; row < endRow; row += 1)
			subjects.add(stringCellAt(textColumn, row));
		if(subjectsMean >= 13) subjectsLevel = 3;
		else if(subjectsMean >= 11) subjectsLevel = 2;
		else if(subjectsMean >= 9) subjectsLevel = 1;

		if(intCellAt(H, 6) != subjectsLevel) {
			problems.add(new Problem(H, 6, "Stufe stimmt nicht mit berechneter Stufe überein (wurde die Exceldatei manipuliert`?)", true));
		}
	}

	private Cell cellAt(int column, int row) {
		return sheet.getRow(row - 1).getCell(column - 1);
	}

	private boolean cellEmpty(int column, int row) {
		Cell cell = cellAt(column, row);
		if(cell.getCellType() == Cell.CELL_TYPE_STRING) {
			String value = cell.getStringCellValue();
			if(value.trim().isEmpty()) return true;
		}
		return cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK;
	}

	private String stringCellAt(int column, int row) {
		Cell cell = cellAt(column, row);
		if (cell.getCellType() == Cell.CELL_TYPE_STRING) return cell.getStringCellValue().trim();
		else {
			int number = intCellAt(column, row);
			if (number == 0) return "";
			else return Integer.toString(number);
		}
	}

	private int intCellAt(int column, int row) {
		Cell cell = cellAt(column, row);
		if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
			try {
				String number = cell.getStringCellValue().trim().replace(',', '.');
				if(number.isEmpty()) return 0;
				return (int) Math.round(Double.parseDouble(number));
			} catch (NumberFormatException e) {
				problems.add(new Problem(row, column, "Zahl erwartet", true));
				return 0;
			}
		}
		return (int) Math.round(cell.getNumericCellValue());
	}

	public String getName() {
		return name;
	}

	public Date getDate() {
		return date;
	}

	public ArrayList<String> getSubjects() {
		return subjects;
	}

	public double getSubjectsMean() {
		return subjectsMean;
	}

	public int getSubjectsLevel() {
		return subjectsLevel;
	}

	public int getProjectLevel() {
		return projectLevel;
	}

	public ArrayList<String> getActivities1() {
		return activities1;
	}

	public ArrayList<String> getActivities2() {
		return activities2;
	}

	public Sheet getSheet() {
		return sheet;
	}

	public int getActivityLevel() {
		return activityLevel;
	}

	public String getProjectString() {
		return projectString;
	}

	public List<Problem> getProblems() {
		return problems;
	}
}
