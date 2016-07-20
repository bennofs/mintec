package mintec;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.swing.SwingWorker;
import javax.swing.table.AbstractTableModel;

import mintec.FileProcessorResult.State;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.PdfCopy;
import com.itextpdf.text.pdf.PdfReader;

class FileTable extends AbstractTableModel {
	private static final long serialVersionUID = -2985648423163667274L;

	private enum Column { STATE, FILE, PROBLEMS }

	private final static String[] columnNames = { "Status", "Datei", "Probleme" };
	private final static Column[] columns = Column.values();
	private final List<FileEntry> files;
	private SwingWorker<Void, FileProcessorResult> worker = null;
	private final byte[] template;
	private final File aggregateDocument;

	FileTable(byte[] template, File aggregateDocument) {
		this.template = template;
		this.aggregateDocument = aggregateDocument;
		this.files = new ArrayList<>();
	}

	private Column toColumn(int index) {
	    return columns[index];
	}

	@Override
	public boolean isCellEditable(int row, int column) { return false; }

	@Override
	public String getColumnName(int column) {
		switch(toColumn(column)) {
			case FILE: return "Datei";
			case STATE: return "Status";
			case PROBLEMS: return "Probleme";
		}
		throw new RuntimeException("Invalid column");
	}

	@Override
	public int getColumnCount() { return columnNames.length; }

	@Override
	public int getRowCount() { return files.size(); }

	@Override
	public Object getValueAt(int row, int column) {
		FileEntry entry = files.get(row);
		switch(toColumn(column)) {
			case FILE: return entry.inputFile.getName();
			case STATE: return entry.getResult().state;
			case PROBLEMS: return entry.getResult().problemsMessage;
		}
		throw new RuntimeException("Invalid column");
	}

	void addEntry(FileEntry entry) {
		files.add(entry);
		fireTableRowsInserted(files.size() - 1, files.size() - 1);
	}

	private class ProcessEntries extends SwingWorker<Void, FileProcessorResult> {
		private FileProcessorResult processEntry(int fileId, FileEntry entry) {
			List<MintReader.Problem> problems = new ArrayList<>();
			Exception exception = null;
			try {
				MintReader reader = new MintReader(new FileInputStream(entry.inputFile));
				problems = reader.getProblems();
				if(new FileProcessorResult(null, problems, fileId).state != State.FAIL) {
					MintWriter writer = new MintWriter(reader, new ByteArrayInputStream(template), new FileOutputStream(entry.outputFile));
					writer.close();
				}
			} catch(Exception exc) {
				exception = exc;
			}
			return new FileProcessorResult(exception, problems, fileId);
		}

		@Override
		protected Void doInBackground() {
			int i = 0;
			for(FileEntry entry : files) {
				publish(processEntry(i++, entry));
				setProgress(i * 90 / files.size());
			}

			Document outDocument = new Document();
			try {
				PdfCopy copy = new PdfCopy(outDocument, new FileOutputStream(aggregateDocument));
				outDocument.open();
				copy.open();
				for(FileEntry entry : files) {
					if(!entry.outputFile.exists()) continue;
					PdfReader reader = new PdfReader(new FileInputStream(entry.outputFile));
					copy.addDocument(reader);
				}
				copy.close();
				outDocument.close();
			} catch (IOException | DocumentException e) {
				System.out.println("Internal error...");
				e.printStackTrace();
				System.exit(1);
			}
			setProgress(100);

			return null;
		}

		@Override
		protected void process(List<FileProcessorResult> results) {
			for(FileProcessorResult result : results) {
				files.get(result.fileId).setResult(result);
				fireTableRowsUpdated(result.fileId, result.fileId);
			}
		}

	}

	SwingWorker<Void, FileProcessorResult> process() {
		if(this.worker != null) {
			this.worker.cancel(true);
			this.worker = null;
		}
		this.worker = new ProcessEntries();
		return this.worker;
	}

}
