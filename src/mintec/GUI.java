package mintec;

import java.awt.BorderLayout;
import java.awt.EventQueue;
import java.awt.FlowLayout;

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JButton;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.SwingWorker;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileFilter;

import org.apache.poi.util.IOUtils;

import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.beans.PropertyChangeEvent;
import java.beans.PropertyChangeListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class GUI extends JFrame {

	private static final long serialVersionUID = -4233818212766136444L;
	private File outputsDirectory;
	private FileTable fileProcessors;

	private JPanel contentPane;
	private JProgressBar progressBar;

	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					JFileChooser outputDirChooser = new JFileChooser();
					outputDirChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
					outputDirChooser.setCurrentDirectory(new File("H:\\"));
					if(outputDirChooser.showDialog(null, "Ausgabeverzeichnis auswählen") != JFileChooser.APPROVE_OPTION) {
						JOptionPane.showMessageDialog(null, "Fehler beim Auswählen des Ausgabeverzeichnisses", "Fehler", JOptionPane.ERROR_MESSAGE);
						System.exit(1);
					}
					GUI frame = new GUI(outputDirChooser.getSelectedFile());
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public GUI(File outputsDirectory) throws IOException {
		org.apache.log4j.BasicConfigurator.configure();
		this.outputsDirectory = outputsDirectory;

		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 450, 300);
		contentPane = new JPanel(new BorderLayout());
		setContentPane(contentPane);

		JPanel controls = new JPanel(new FlowLayout(FlowLayout.LEADING));
		contentPane.add(controls, BorderLayout.SOUTH);

		JButton btnHinzufgen = new JButton("Hinzufügen");
		btnHinzufgen.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser inputFile = new JFileChooser();
				ExcelFilter excel = new ExcelFilter();
				inputFile.setFileFilter(excel);
				inputFile.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
				if(inputFile.showOpenDialog(GUI.this) == JFileChooser.APPROVE_OPTION) {
					File result = inputFile.getSelectedFile();

					File[] files = { result };
					if(result.isDirectory()) files = result.listFiles(excel);
					for(File file : files) {
						if(file.isDirectory()) continue;
						File out = new File(GUI.this.outputsDirectory.getAbsolutePath() + "\\" + file.getName() + ".pdf");
						GUI.this.fileProcessors.addEntry(new FileEntry(file, out));
					}
				};

			}
		});
		controls.add(btnHinzufgen);

		JButton btnProcess = new JButton("Zertifikate erstellen");
		btnProcess.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ignored) {
				SwingWorker<Void, FileProcessorResult> worker = GUI.this.fileProcessors.process();
				worker.addPropertyChangeListener(new PropertyChangeListener() {
					public void propertyChange(PropertyChangeEvent event) {
						if("progress".equals(event.getPropertyName())) {
							GUI.this.progressBar.setValue((Integer)event.getNewValue());
						}
					}
				});
				worker.execute();
			}
		});
		controls.add(btnProcess);

		JPanel main = new JPanel(new BorderLayout());
		main.setBorder(new EmptyBorder(10,10,10,10));
		contentPane.add(main, BorderLayout.CENTER);

		JLabel lblAusgabeverzeichnis = new JLabel("Ausgabeverzeichnis: " + outputsDirectory.getAbsolutePath());
		main.add(lblAusgabeverzeichnis, BorderLayout.NORTH);

		progressBar = new JProgressBar(0, 100);
		main.add(progressBar, BorderLayout.SOUTH);


		final byte[] template;
		File path = new File(GUI.class.getProtectionDomain().getCodeSource().getLocation().getPath());
		File folder = path.getParentFile();
		template = IOUtils.toByteArray(new FileInputStream(new File(folder.getAbsolutePath() + "\\template.pdf")));
		fileProcessors = new FileTable(template, new File(outputsDirectory.getAbsolutePath() + "\\all.pdf"));
		JTable fileProcessorTable = new JTable(fileProcessors);
		main.add(new JScrollPane(fileProcessorTable), BorderLayout.CENTER);
	}
}

class ExcelFilter extends FileFilter implements java.io.FileFilter {
	@Override
	public boolean accept(File pathname) {
		return pathname.isDirectory() || pathname.getAbsolutePath().endsWith(".xls") || pathname.getAbsolutePath().endsWith(".xlsx");
	}

	@Override
	public String getDescription() {
		return "MINT-EC Antrag oder Verzeichnis mit Anträgen";
	}
}
