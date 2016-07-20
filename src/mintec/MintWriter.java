package mintec;

import java.io.IOException;
import java.io.OutputStream;
import java.io.InputStream;
import java.math.RoundingMode;
import java.text.SimpleDateFormat;

import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.AcroFields;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;

import java.text.DecimalFormat;
import java.util.Date;
import java.util.Locale;

/**
 * Diese Klasse ist dafür verantwortlich, aus den von {@link MintReader} eingelesenen Daten ein MINTEC-Zertifikat zu
 * generieren.
 *
 * Dazu wird eine PDF-Vorlage, welche vom MINT-EC Verein zur Verfügung gestellt wird, verwendet. Diese PDF-Datei besitzt
 * bereits fertige Formularfelder, die bestimmte Textbereiche des Zertifikats kennzeichnen.
 */
public class MintWriter {
    /** Die Vorlage als PDF-Dokument. */
    private PdfReader doc;

    /**
     * Ein Stamper zur Generierung des PDF-Dokuments für das Zertifikat.
     *
     * Der Stamper erlaubt es, einzelne Formularfelder in dem PDF-Dokument auszufüllen und so ein neues PDF zu
     * generieren.
     */
    private PdfStamper stamper;

    /**
     * Ein Objekt, welches die Formularfelder des PDF-Dokuments repräsentiert.
     */
    private AcroFields fields;

    /**
     * Erstellt ein neues MINT-EC Zertifikat, in dem die Daten aus einem {@link MintReader} in ein PDF-Formular
     * übertragen werden.
     *
     * @param reader Der MintReader, aus welchem die Daten gelesen werden.
     * @param template Die Vorlage für das PDF-Formular.
     * @param result OutputStream, in den das Ergebnis geschrieben werden soll.
     * @throws IOException Wenn Fehler beim Schreiben auftreten.
     * @throws DocumentException Wenn Fehler beim Schreiben oder Lesen des PDF-Dokuments auftreten.
     */
	MintWriter(MintReader reader, InputStream template, OutputStream result) throws IOException, DocumentException {
        // Lese die Vorlage und initialisiere darauf aufbauend den Stamper, der für die Generierung des Zertifikats
        // verantwortlich ist.
		doc = new PdfReader(template);
        stamper = new PdfStamper(doc, result);
		fields = stamper.getAcroFields();

        // Schreibe die Daten in das PDF-Dokument
		setField("Vor- und Nachname", reader.getName());
        setField("Schulbezeichnung", "am Martin-Andersen-Nexö Gymnasium Dresden");

        Date d = reader.getDate();
		String dateString = d != null ? new SimpleDateFormat("dd. MMMM yyyy", Locale.GERMAN).format(d) : "N/A";
		setField("geboren am Tag / Monat / Jahr", "geboren am " + dateString);

        // Die erreichte Stufe ist der gerundete Durchschnitt aus den Stufen in den drei einzelnen Bereichen
		long total = Math.round(((double)reader.getSubjectsLevel() + (double)reader.getProjectLevel() + (double)reader.getActivityLevel()) / 3.0);
		String[] names = { "mit Erfolg", "mit besonderem Erfolg", "mit Auszeichnung" }; // Die Bezeichnungen der Stufen
		setField("Gesamteinstufung", names[(int)total-1]);

        // Erstelle den Text für den Abschnitt I - "Fachliche Kompetenz".
        // Das Ergebnis sieht wie folgt aus:
        //
        // [Fach1]
        // <Leerzeile>
        // [Fach2]
        // <Leerzeile>
        // [Fach3] (nur wenn Variante 2 gewählt wurde, bei der 3 Fächer angegeben werden müssen.)
        // <Leerzeile>
        // Durchschnittliche Note über alle [zwei oder drei] Fächer gemittelt: [Durchschnitt] Punkte
        String subjectsString = "";
        for(String subject : reader.getSubjects()) subjectsString += subject + "\n\n";
		String numberWord = reader.getSubjects().size() == 3 ? "drei" : "zwei";
		DecimalFormat formatter = new DecimalFormat("0.0");
		formatter.setRoundingMode(RoundingMode.HALF_UP); // Der Standard-Rundungsmodus ist HALF_EVEN, was falsch wäre.
		String pointsString = formatter.format(reader.getSubjectsMean());
		subjectsString += "\nDurchschnittliche Note über alle " + numberWord + " Fächer gemittelt: " + pointsString + " Punkte";
        setField("Fachliche Kompetenz", subjectsString);

        // Der Text für den Abschnitt II - "Fachwissenschaftliches Arbeiten" wurde schon vom Reader erstellt, sodass wir
        // einfach den fertigen Text verwenden können.
        setField("Fachwissenschaftliches Arbeiten", reader.getProjectString());


        // Für Abschnitt III - "Zusätzliche MINT-Aktivitäten" generieren wir einfach eine Auflistung aller angegebenen
        // Aktivitäten, wie folgt:
        //
        // In der Sekundarstufe 1:
        // <Leerzeile>
        // [Liste der Namen aller Aktivitäten für Sekundarstufe 1, eine Name pro Zeile]
        // <Leerzeile>
        // In der Sekundarstufe 2:
        // <Leerzeile>
        // [Liste der Namen aller Aktivitäten für Sekundarstufe 2, eine Name pro Zeile]
        String extraString = "In der Sekundarstufe 1:\n\n";
		for(String item : reader.getActivities1()) extraString += item + "\n";
		extraString += "\nIn der Sekundarstufe 2:\n\n";
		for(String item : reader.getActivities2()) extraString += item + "\n";
		setField("Zusätzliche MINT-Aktivitäten", extraString);
	}

    /**
     * Schließt die Generierung eines MINT-EC Zertifikats ab.
     *
     * Diese Method sollte aufgerufen werden, um das fertige MINT-EC Zertifikat zu schreiben.
     *
     * @throws IOException Wenn Fehler beim Schreiben auftreten (IO).
     * @throws DocumentException Wenn PDF-Fehler auftreten.
     */
	void close() throws IOException, DocumentException {
		stamper.setFormFlattening(true);
		stamper.close();
		doc.close();
	}

    /**
     * Setzt ein PDF-Formularfeld in der Vorlage auf einen bestimmte Wert.
     *
     * @param id Name des Formularfeldes, dessen Wert gesetzt werden soll.
     * @param data Der Text, der in das Formularfeld geschrieben werden soll.
     * @throws IOException Wenn Fehler beim Schreiben auftreten (IO).
     * @throws DocumentException Wenn PDF-Fehler auftreten.
     */
	private void setField(String id, String data) throws IOException, DocumentException {
		fields.setField(id, data);
	}
}
