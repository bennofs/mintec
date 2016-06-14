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

public class MintWriter {
	PdfReader doc;
	PdfStamper stamper;
	AcroFields fields;

	MintWriter(MintReader reader, InputStream template, OutputStream result) throws IOException, DocumentException {
		doc = new PdfReader(template);

		try {
			stamper = new PdfStamper(doc, result);
		} catch(DocumentException e) {
			throw new RuntimeException(e);
		}

		fields = stamper.getAcroFields();
		setField("Vor- und Nachname", reader.getName());
		Date d = reader.getDate();

		String dateString = d != null ? new SimpleDateFormat("dd. MMMM yyyy").format(d) : "N/A";
		setField("geboren am Tag / Monat / Jahr", "geboren am " + dateString);
		setField("Schulbezeichnung", "am Martin-Andersen-Nexö Gymnasium Dresden");

		long total = Math.round(((double)reader.getSubjectsLevel() + (double)reader.getProjectLevel() + (double)reader.getActivityLevel()) / 3.0);
		String[] names = { "mit Erfolg", "mit besonderem Erfolg", "mit Auszeichnung" };
		setField("Gesamteinstufung", names[(int)total-1]);

		String subjectsString = "";
		String numberWord = reader.getSubjects().size() == 3 ? "drei" : "zwei";
		DecimalFormat formatter = new DecimalFormat("0.0");
		formatter.setRoundingMode(RoundingMode.HALF_UP);
		String pointsString = formatter.format(reader.getSubjectsMean());
		for(String subject : reader.getSubjects()) subjectsString += subject + "\n\n";
		subjectsString += "\nDurchschnittliche Note über alle " + numberWord + " Fächer gemittelt: " + pointsString + " Punkte";
		String extraString = "In der Sekundarstufe 1:\n\n";
		for(String item : reader.getActivities1()) extraString += item + "\n";
		extraString += "\nIn der Sekundarstufe 2:\n\n";
		for(String item : reader.getActivities2()) extraString += item + "\n";
		setField("Fachliche Kompetenz", subjectsString);
		setField("Fachwissenschaftliches Arbeiten", reader.getProjectString());;
		setField("Zusätzliche MINT-Aktivitäten", extraString);
	}

	void close() throws IOException, DocumentException {
		stamper.setFormFlattening(true);
		stamper.close();
		doc.close();
	}

	private void setField(String id, String data) throws IOException, DocumentException {
		fields.setField(id, data);
	}
}
