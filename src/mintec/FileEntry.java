package mintec;

import java.io.File;

public class FileEntry {
	public final File inputFile;
	public final File outputFile;

	private FileProcessorResult result;

	public FileEntry(File in, File out) {
		super();
		this.inputFile = in;
		this.outputFile = out;
		this.result = new FileProcessorResult();
	}

	public FileProcessorResult getResult() {
		return this.result;
	}

	public void setResult(FileProcessorResult result) {
		this.result = result;
	}
}
