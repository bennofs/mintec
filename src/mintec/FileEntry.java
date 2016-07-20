package mintec;

import java.io.File;

class FileEntry {
	final File inputFile;
	final File outputFile;

	private FileProcessorResult result;

	FileEntry(File in, File out) {
		super();
		this.inputFile = in;
		this.outputFile = out;
		this.result = new FileProcessorResult();
	}

	FileProcessorResult getResult() {
		return this.result;
	}

	void setResult(FileProcessorResult result) {
		this.result = result;
	}
}
