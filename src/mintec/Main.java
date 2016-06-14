package mintec;

import java.io.File;
import java.io.FileFilter;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfCopy;
import com.itextpdf.text.pdf.PdfReader;


public class Main {

	public static void main(String[] args) throws Exception {
		org.apache.log4j.BasicConfigurator.configure();

		File inputsDirectory = new File("H:\\mintec\\inputs");
		FileFilter excel = new FileFilter() {
			@Override
			public boolean accept(File pathname) {
				return pathname.getAbsolutePath().endsWith(".xls") || pathname.getAbsolutePath().endsWith(".xlsx");
			}
		};
		File[] inputs = inputsDirectory.listFiles(excel);

		File outputsDirectory = new File("H:\\mintec\\outputs");
		outputsDirectory.mkdir();

		System.out.println("== Discovered " + inputs.length + " files");
		for(File in : inputs) {
			File out = new File(outputsDirectory.getAbsolutePath() + "\\" + in.getName() + ".pdf");
			if(out.exists()) {
				System.out.println("= Skipping: " + in.getName());
				continue;
			}
			System.out.println("= Processing: " + in.getName());

			MintReader reader = new MintReader(new FileInputStream(in));
			MintWriter writer = new MintWriter(reader, new FileInputStream("H:\\mintec\\vorlage.pdf"), new FileOutputStream(out.getAbsolutePath()));
			writer.close();
		}
		System.out.println("== Merging PDFs");
		File out = new File(outputsDirectory.getAbsolutePath() + "\\all.pdf");
		Document outDocument = new Document();
		PdfCopy copy = new PdfCopy(outDocument, new FileOutputStream(out));
		outDocument.open();
		copy.open();
		for(File in : inputs) {
			String pdf = outputsDirectory.getAbsolutePath() + "\\" + in.getName() + ".pdf";
			System.out.println("= Adding: " + pdf);
			PdfReader reader = new PdfReader(pdf);
			copy.addDocument(reader);
		}
		copy.close();
		outDocument.close();
	}

}
