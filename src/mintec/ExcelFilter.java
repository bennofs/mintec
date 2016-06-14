package mintec;

import javax.swing.filechooser.FileFilter;
import java.io.File;

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
