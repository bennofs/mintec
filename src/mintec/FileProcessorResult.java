package mintec;

import java.util.List;

public class FileProcessorResult {

	enum State {
		OK,
		FAIL,
		WARN
	}

	public final String problemsMessage;
	public final State state;
	public final int fileId;

	public FileProcessorResult() {
		this.state = null;
		this.problemsMessage = "";
		this.fileId = -1;
	}

	public FileProcessorResult(Exception exception, List<MintReader.Problem> problems, int fileId) {
		State state = State.OK;

		StringBuilder errors = new StringBuilder();
		StringBuilder warnings = new StringBuilder();
		if(exception != null) errors.append("Ein-/Ausgabefehler: " + exception.getLocalizedMessage() + "\n");
		for(MintReader.Problem problem : problems) {
			String msg = "Zelle " + problem.column + problem.row + ": " + problem.text + "\n";
			if(problem.fatal) {
				state = State.FAIL;
				errors.append(msg);
			} else {
				state = State.WARN;
				warnings.append(msg);
			}
		}

		this.state = state;
		this.problemsMessage = errors.append("\n").append(warnings).toString();
		this.fileId = fileId;
	}
}