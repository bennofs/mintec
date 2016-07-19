package mintec;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * Klasse für das Einlesen der Daten eines MINTEC-Antrags aus einer Exceldatei.
 *
 * Diese Klasse ließt die Excel-Datei eines MINTEC-Antrags ein und stellt die darin enthalten
 * Daten bereit. Beim Einlesen wird der Antrag gleichzeitig auf Fehler überprüft (sind alle
 * benötigten Felder ausgefüllt? usw).
 */
public class MintReader {
  /**
   * Konstanten für die Indices der Excel-Spalten.
   *
   * Die verwendete Bibliothek zum Einlesen von Excel-Dateien verwendet sowohl für Spalten als auch
   * für Zeilen numerische Indices. Zur besseren Lesbarkeit werden hier für die entsprechenden Namen
   * der Excel-Spalten Konstanten definiert, die später anstelle der Indices für die Spalten
   * verwendet werden.
   */
  @SuppressWarnings("unused")
  private final int A = 1, B = 2, C = 3, D = 4, E = 5, F = 6, G = 7, H = 8;

  /** Die Excel-Tabelle des Antrags. */
  private Sheet sheet;

  /**
   * Eine Liste von Problemen (Fehler oder Warnung), die während des Einlesens festgestellt wurden.
   *
   * Die Verwendung einer Liste hier erlaubt es, möglichst alle Fehler auf einmal zu finden und
   * nicht nach einem Fehler direkt abzubrechen (was bei der Verwendung von Exceptions der Fall
   * wäre). Somit können die meisten Fehler direkt bei der ersten Ausführung des Programms gefunden
   * werden, und es muss nicht nach jedem Fehler das Programm erneut ausgeführt werden, nur um den
   * nächsten Fehler zu finden.
   *
   * @return Die Liste festgestellten Probleme.
   */
  @Getter private List<Problem> problems;

  /**
   * Der Name des Schülers.
   *
   * @return Name des Schülers.
   */

  @Getter private String name;

  /**
   * Das Geburtsdatum des Schülers.
   *
   * @return Geburtsdatum des Schülers.
   */
  @Getter private Date date;

  /**
   * Die Namen der Fächer, welche in Sektion 1 angegeben wurden.
   *
   * Dieses Feld enthält entweder:
   * <ul>
   * <li> zwei Fächer, für die Option "Zwei Abiturfächer auf erhöhtem Niveau"
   * <li> drei Fächer, für die Option "Ein Abiturfach auf erhötem Niveau und zwei weitere, in der
   *      Qualifikationsphase durchgängig belegte Fächer:"
   * </ul>
   *
   * Wenn beide Varianten vollständig ausgefüllt wurden, wird die erste Variante bevorzugt.
   *
   * @return Name der in Sektion 1 angegebenen Fächer.
   */
  @Getter private ArrayList<String> subjects;

  /**
   * Der Durchschnitt der Noten (alle 4 Halbjahre) aller in {@link #subjects} genannten Fächer.
   *
   * @return Durchschnitt der Noten der angegebenen Fächer, in Punkten.
   */
  @Getter private double subjectsMean;

  /**
   * Die Stufe für Abschnitt I - entweder 0, 1, 2 oder 3.
   *
   * @return Stufe für Abschnitt I.
   */
  @Getter private int subjectsLevel;

  /**
   * Ein Text, der die Arbeit aus dem Bereich "Fachwissenschaftliches Arbeiten" beschreibt.
   *
   * Für eine BeLL würde dieses Feld beispielsweise folgenden Text enthalten:
   * <pre>
   * Besondere Lernleistung
   *
   * Fach: [Fach der BeLL]
   *
   * Thema:
   * [Thema der BeLL]
   *
   * Note: [Note der BeLL]
   * </pre>
   *
   * Der Text in diesem Feld wird direkt für den Abschnitt "Fachwissenschaftliches Arbeiten" des
   * Zertifikats verwendet.
   *
   * @return Text, der Abschnitt II beschreibt.
   */
  @Getter private String projectString = "";

  /**
   * Die Stufe für Abschnitt II - entweder 0, 1, 2 oder 3.
   *
   * @return Stufe für Abschnitt II.
   */
  @Getter private int projectLevel;

  /**
   * Liste der zusätzlichen MINT-Aktivitäten aus der Sekundarstufe 1.
   *
   * Diese Liste enthält einfach nur die Namen aller MINT-Aktivitäten aus dem Abschnitt
   * III - "Zusätzliche MINT-Aktivitäten" für die Sekundarstufe 1.
   *
   * @return Namen der zusätlichen MINT-Aktivitäten für die Sekundarstufe 1
   */
  @Getter private ArrayList<String> activities1;

  /**
   * Liste der zusätzlichen MINT-Aktivitäten aus der Sekundarstufe 2.
   *
   * Diese Liste enthält einfach nur die Namen aller MINT-Aktivitäten aus dem Abschnitt
   * III - "Zusätzliche MINT-Aktivitäten" für die Sekundarstufe 2.
   *
   * @return Namen der zusätlichen MINT-Aktivitäten für die Sekundarstufe 1
   */
  @Getter private ArrayList<String> activities2;

  /**
   * Die Stufe für Abschnitt III - entweder 0, 1, 2 oder 3.
   * @return Stufe für Abschnitt III.
   */
  @Getter private int activityLevel;

  /** Klasse für festgestellte Probleme (entweder Warnung oder Fehler) beim Einlesen. */
  @AllArgsConstructor
  public class Problem {
    /** Die Zeile der Zelle, deren Wert fehlerhaft ist und somit das Problem verursacht. */
    public final int row;

    /** Der Name der Spalte der Zelle, deren Wert fehlerhaft ist und somit das Problem verursacht. */
    public final String column;

    /** Ein Text, der das Problem möglichst genau beschreibt. */
    public final String text;

    /** Legt fest, ob das Problem die Generierung eines Zertifikates unterbinden soll.
     *
     * In einigen Fällen ist es möglich, trotzdem ein Zertifikat zu erstellen, obwohl ein Problem
     * auftrat. Das kann zum Beispiel vorkommen, wenn das Problem automatisch korrigiert werden
     * konnte. In diesem Fall sollte dieses Feld auf false gesetzt werden, dann wird trotz des
     * Problems ein Zertifikat generiert und lediglich eine Warnung ausgegeben.
     */
    public final Boolean fatal;

    /** Erstellt ein neues Problem.
     *
     * @param row Zeilenindex der Zelle, deren Wert fehlerhaft ist und somit das Problem verursacht.
     * @param column Spaltenindex der Zelle, deren Wert fehlerhaft ist und somit das Problem verursacht.
     * @param text Ein Text, der das Problem möglichst genau beschreibt.
     * @param fatal Legt fest, ob dies ein kritisches Problem ist. Siehe {@link #fatal}.
     */
    public Problem(int row, int column, String text, Boolean fatal) {
      super();
      String columnLetters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
      this.row = row;
      this.column = "" + columnLetters.charAt(column - 1);
      this.text = text;
      this.fatal = fatal;
    }

  }

  /**
   * Dieser Konstruktor liest die Daten aus einem MINT-Zertifikat Antrag und setzt die
   * entsprechenden Membervariablen.
   *
   * @param file InputStream für die Excel-Datei, die die Daten dieses Antrags enthält.
   * @throws IOException Wenn das Lesen aus der Datei fehlschlägt.
   * @throws InvalidFormatException Wenn eine Zelle ein ungültiges Format hat.
   */
  MintReader(FileInputStream file) throws InvalidFormatException, IOException {
    sheet = WorkbookFactory.create(file).getSheetAt(0);
    problems = new ArrayList<>();

    // Es gibt mehrere Versionen des Formulars, die inkompatibel sind. Um sicher zu gehen, dass
    // wir keine alte Version bekommen, wird hier überprüft, dass die Formularversion diejenige ist,
    // die wir erwarten.
    {
      final String expectedVersion = "1.0.0";
      final String readVersion = stringCellAt(A,3);
      if(!expectedVersion.equals(readVersion)) {
        problems.add(new Problem(A, 3, "Inkompatible Formularversion: " + "Version ist " + readVersion + ", erwarte " + expectedVersion, true));
        return;
      }
    }

    // Einlesen der Daten zum Antragssteller (Name, Geburtsdatum)
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

    // Einlesen der Daten aus den verschiedenen Abschnitten
    final int textColumn = B;
    final int gradeColumn = D;
    final int levelColumn = H;

    extractSubjects(textColumn, gradeColumn);
    extractProjects(textColumn, gradeColumn);
    extractActivities(textColumn, levelColumn);
  }

  /** Diese Methode liest die Daten aus dem Abschnitt I - Fachliche Kompetenz ein.
   *
   * @param textColumn Spalte, in der die Fächernamen stehen.
   * @param gradeColumn Spalte, welche die Noten in dem jeweiligen Fach enthält.
   */
  private void extractSubjects(final int textColumn, final int gradeColumn) {
    // Die Spalte/Zeile der Zelle, die den Durchschnitt der Noten in den Fächern enthält
    final int meanColumn = E;
    final int meanRow = 7;

    // Für Abschnitt I existieren zwei Varianten:
    // - "Zwei Abiturfächer auf erhöhtem Niveau", oder
    // - "Ein Abiturfach auf erhöhtem Niveau und zwei weitere, in der Qualifikationsphase durchgängig belegte Fächer"
    // Die erste Variante ist im folgenden Code mit "twoSubjects" bezeichnet (da zwei Fächer
    // benötigt werden), die zweite mit "threeSubjects".

    // Zeilen für die Fächer bei Variante 1
    final int[] twoSubjectsRows = {7, 8};

    // Zeilen für die Fächer bei Variante 2
    final int[] threeSubjectsRows = {10,11,12};

    // Wir müssen entscheiden, welche der beiden Varianten für Abschnitt I ausgefüllt ist bzw.
    // welche verwendet werden soll. Damit das Programm in möglichst vielen Fällen funktioniert,
    // möchten wir die Variant verwenden, welche vollständig ausgefüllt ist, selbst wenn die andere
    // Variante teilweise auch ausgefüllt ist (In diesem Fall wird eine Warnung generiert).
    //
    // Es reicht also nicht, allein zu prüfen, ob die erste Zeile einer Variante vollständig
    // ausgefüllt ist, denn es kann immer noch passieren, das eine andere Zeile nicht vollständig
    // ausgefüllt ist und wir daher trotzdem die andere Variante auswählen müssen.
    //
    // Stattdessen zählen wir jedes ausgefüllte Feld einer Variante. Wenn die Anzahl der ausgefüllten
    // Zellen gleich der Anzahl der Zellen ist, die für diese Variante ausgefüllt werden müssen,
    // dann wissen wir, dass diese Variante vollständig ausgefüllt ist.
    //
    // Wir können also nach diesem Test einfach entscheiden, ob und wenn ja welche der beiden
    // Varianten vollständig ausgefüllt ist und dann diese auslesen.

    // Zuerst werden die ausgefüllten Felder für jede der beiden Varianten gezählt.
    int twoSubjectsFilledFields = 0, threeSubjectsFilledFields = 0;
    for(int column : new int[]{textColumn, gradeColumn}) {
      for (int row : twoSubjectsRows) {
        if (!cellEmpty(column, row)) twoSubjectsFilledFields += 1;
      }
      for (int row : threeSubjectsRows) {
        if (!cellEmpty(column, row)) threeSubjectsFilledFields += 1;
      }
    }

    // Danach können wir uns für eine der beiden Varianten entscheiden.
    subjects = new ArrayList<>();
    if(twoSubjectsFilledFields == 4) {
      for(int row : twoSubjectsRows) subjects.add(stringCellAt(textColumn, row));
      if(threeSubjectsFilledFields > 0) {
        problems.add(new Problem(
          threeSubjectsRows[0],
          textColumn,
          "Ignoriere überflüssige Daten für Variante „Ein Abiturfach und zwei weitere Fächer”",
          false
        ));
      }
    } else if(threeSubjectsFilledFields == 6) {
      for(int row : threeSubjectsRows) subjects.add(stringCellAt(textColumn, row));
      if (twoSubjectsFilledFields > 0)
        problems.add(new Problem(
          twoSubjectsRows[0],
          textColumn,
          "Ignoriere überflüssige Daten für Variante „Zwei Abiturfächer auf erhötem Niveau”",
          false
        ));
    } else {
      problems.add(new Problem(A, 4,
        "Weder zwei Abiturfächer auf erhöhtem Niveau noch ein Abiturfach auf erhöhtem Niveau und zwei weitere Fächer vollständig ausgefüllt.",
        true
      ));
    }

    subjectsMean = cellAt(meanColumn, meanRow).getNumericCellValue();
    if(subjectsMean >= 13) subjectsLevel = 3;
    else if(subjectsMean >= 11) subjectsLevel = 2;
    else if(subjectsMean >= 9) subjectsLevel = 1;

    if(intCellAt(H, 6) != subjectsLevel) {
      problems.add(new Problem(H, 6,
        "Stufe stimmt nicht mit berechneter Stufe überein (wurde die Exceldatei manipuliert?)",
        true
      ));
    }
  }

  /** Liest die Daten aus Abschnitt II - Fachwissenschaftliches Arbeiten aus
   *
   * @param textColumn Spalte, die Text wie Name einer fachwissenschaftlichen Arbeit oder die
   *                   Beschreibung einer fachwissenschaftlichen Arbeit enthält.
   * @param gradeColumn Spalte, welche die Bewertung einer fachwissenschaftlichen Arbeit enthält.
   */
  private void extractProjects(final int textColumn, final int gradeColumn) {
    // Spalte für die aus einer einzigen fachwissenschaftlichen Arbeit ermittelte Stufe.
    final int subLevelColumn = E;

    // Es gibt vier Varianten für eine fachwissenschaftliche Arbeit:
    // - Wissenschaftspropädeutisches Fach (A)
    // - Fachwissenschaftliche Arbeit (B)
    // - Besondere Lernleistung (C)
    // - Jugend forscht-Wettbewerb / vergleichbarer Wettbewerb (D)

    // Diese Variablen kennzeichen die erste Zeile, die Daten der jeweiligen Variante enthält.
    final int rowA = 15;
    final int rowB = 17;
    final int rowC = 20;
    final int rowD = 23;

    // Als nächstes werden die von der Excel-Tabelle berechneten Stufen für die einzelnen
    // Varianten ausgelesen. Das ist notwendig, um die richtige Variante für das Zertifikat
    // auszuwählen: wenn mehrere Varianten ausgefüllt sind, dann bevorzugen wir diejenige, die
    // die höchste Stufe hat.

    final int levelA = intCellAt(subLevelColumn, rowA);
    final int levelB = intCellAt(subLevelColumn, rowB);
    final int levelC = intCellAt(subLevelColumn, rowC);
    int levelD = intCellAt(subLevelColumn, rowD);

    if(levelD > 3) {
      problems.add(new Problem(rowD, subLevelColumn, "Stufe des Jugend forscht-Wettbewerbs größer als 3", true));
      return;
    }

    // Die in diesem Abschnitt erreichte Stufe ist das Maximum aller in den einzelnen Varianten
    // erreichten Stufen, da wir nur die beste Variante werten.
    projectLevel = Math.max(levelA, Math.max(levelB, Math.max(levelC, levelD)));

    // Nun können wir die Variante auswählen, die wir für das Zertifikat verwenden wollen:
    // das ist einfach die Variante, mit der die beste Stufe (projectLevel) erreicht wird.

    // Erste Zeile der Daten der besten Variante (enthält Fach des Projekts bzw. Name des Wettbewerbs)
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
      throw new AssertionError("this should be impossible");
    }

    if(stringCellAt(textColumn, nameRow).isEmpty()) {
      problems.add(new Problem(nameRow, textColumn, "Fehlende Eingabe", true));
    }
    projectString += stringCellAt(textColumn, nameRow);

    // Alle Varianten außer A haben nach der Name-Zeile eine Zeile mit dem Thema der Arbeit
    if(nameRow != rowA) {
      if(stringCellAt(textColumn, nameRow + 1).isEmpty()) {
        problems.add(new Problem(nameRow + 1, textColumn, "Fehlende Eingabe", true));
      }
      projectString += "\n\n" + "Thema:\n" + stringCellAt(textColumn, nameRow + 1);
    }

    // Bei Jugend forscht gibt es eine "Ergebnis" Zelle, die statt der Note verwendet wird
    if(nameRow == rowD) {
      if(stringCellAt(textColumn, nameRow + 2).isEmpty()) {
        problems.add(new Problem(nameRow + 2, textColumn, "Fehlende Eingabe", true));
      }
      projectString += "\n\n" + stringCellAt(textColumn, nameRow + 2);
    } else
      projectString += "\n\n" + "Note: " + Integer.toString(intCellAt(gradeColumn, nameRow));
  }

  /**
   * Liest die Daten aus dem Abschnitt "Zusätzliche MINT-Aktivitäten" ein.
   *
   * @param textColumn Spalte, welche die Namen der Aktivitäten enthält.
   * @param levelColumn Spalte, welche die Stufe für diesen Abschnitt enthält.
   */
  private void extractActivities(final int textColumn, final int levelColumn) {
    // Start/End-Zeile für Aktivitäten aus der Sekundarstufe 1
    final int activities1startRow = 28;
    final int activities1endRow = activities1startRow + 20;

    // Start/End-Zeile für Aktivitäten aus der Sekundarstufe 2
    final int activities2startRow = 49;
    final int activities2endRow = activities2startRow + 20;

    // Letzte Zeile dieses Abschnitts, enthält Stufe des Abschnitts
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

  /**
   * Findet die Zelle mit gegebener Spalte und Zeile.
   *
   * Für die Indices der Spalten A bis H sind in dieser Klasse Konstanten definiert,
   * sodass z.B. cellAt(A,1) geschrieben werden kann um die Zelle A1 zu bekommen.
   *
   * @param column Index der Spalte (1-basiert, 1 ist die erste Spalte).
   * @param row Index der Zeile (1-basiert, 1 ist die erste Zeile)
   * @return Das Cell-Objekt für diese Zelle.
   */
  private Cell cellAt(int column, int row) {
    return sheet.getRow(row - 1).getCell(column - 1);
  }

  /**
   * Überprüft, ob eine bestimmte Zelle leer ist (nur Leerzeichen enthält).
   *
   * @param column Index der Spalte (1-basiert, 1 ist die erste Spalte). Siehe {@link #cellAt}.
   * @param row Index der Zeile (1-basiert, 1 ist die erste Zeile). Siehe {@link #cellAt}.
   * @return true wenn die Zelle leer ist, sonst false.
   */
  private boolean cellEmpty(int column, int row) {
    Cell cell = cellAt(column, row);
    if(cell.getCellType() == Cell.CELL_TYPE_STRING) {
      String value = cell.getStringCellValue();
      if(value.trim().isEmpty()) return true;
    }
    return cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK;
  }

  /**
   * Liest den Inhalt einer Zelle als Text.
   *
   * Die Funktion versucht, Fehler möglichst automatisch zu beheben. Falls die Zelle bswp. eine
   * Zahl enthält, wird kein Fehler generiert, sondern die Zahl als Text zurückgegeben.
   *
   * @param column Index der Spalte (1-basiert, 1 ist die erste Spalte). Siehe {@link #cellAt}.
   * @param row Index der Zeile (1-basiert, 1 ist die erste Zeile). Siehe {@link #cellAt}.
   * @return Text-Repräsentation des Inhalts einer Zelle.
   */
  private String stringCellAt(int column, int row) {
    Cell cell = cellAt(column, row);
    if (cell.getCellType() == Cell.CELL_TYPE_STRING) return cell.getStringCellValue().trim();
    else {
      int number = intCellAt(column, row);
      if (number == 0) return "";
      else return Integer.toString(number);
    }
  }

  /**
   * Liest den Inhalt einer Zelle als Zahl.
   *
   * Die Funktion versucht, Fehler möglichst intelligent zu beheben. Falls die Zelle bspw. eine
   * Text-Zelle ist, wird versucht, den Inhalt der Text-Zelle als Zahl zu interpretieren. Dazu
   * versteht die Funktion sowohl . als auch , für Zahlen mit Nachkommastellen.
   *
   * @param column Index der Spalte (1-basiert, 1 ist die erste Spalte). Siehe {@link #cellAt}.
   * @param row Index der Zeile (1-basiert, 1 ist die erste Zeile). Siehe {@link #cellAt}.
   * @return Den Inhalt der Zelle als Zahl, auf ganze Zahlen gerundet.
   */
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
}
