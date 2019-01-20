package fe.form.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JTextPane;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class So1Cua {
	private int approve = 0;
	private File file;
	public static int currentNumber = 0;
	private int checkRepetition = 0;

	public int getCheckRepe() {
		return this.checkRepetition;
	}

	public int getApprove() {
		return this.approve;
	}

	public File getFile() {
		return this.file;
	}

	public void setFile(File file) {
		this.file = file;
	}

	public void setCheckRepe(int num) {
		this.checkRepetition = num;
	}

	public void chooseFile(JFrame frame) {
		JFileChooser chooser = new JFileChooser("./");
		FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel files only", "xlsx", "xls");
		chooser.addChoosableFileFilter(filter);
		chooser.setAcceptAllFileFilterUsed(false);
		
		int returnVal = chooser.showOpenDialog(null);
		if (returnVal == JFileChooser.APPROVE_OPTION) {
			System.out.println("You chose to open this file: " + chooser.getSelectedFile().getName());
			this.file = chooser.getSelectedFile();
			
			FileInputStream inputStream = null;
			try {
				inputStream = new FileInputStream(this.file);
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			}
			XSSFWorkbook workbook = null;
			try {
				workbook = new XSSFWorkbook(inputStream);
			} catch (IOException e) {
				e.printStackTrace();
			}
			try {
				inputStream.close();
			} catch (IOException e1) {
				e1.printStackTrace();
			}
			FileOutputStream out = null;
			try {
				out = new FileOutputStream(this.file);
			} catch (FileNotFoundException e) {
				this.file = null;
				JOptionPane.showMessageDialog(null, "Tắt Sổ 1 Cửa rồi chọn lại!");
			}
			try {
				workbook.write(out);
				out.close();
			} catch (IOException e) {
				frame.dispose();
				e.printStackTrace();
			}
		} else {

		}
	}

	public int checkClosed() {
		try {
			FileWriter fw = new FileWriter(this.file);
		} catch (IOException e) {
			this.file = null;
			return 0;
		}
		return 1;
	}

	public void updateSo1Cua(Form form, JTextPane so1cuathongbao) throws IOException {
		Cell cell;
		Row row;
		CellStyle style;

		FileInputStream inputStream = new FileInputStream(this.file);
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet sheet = workbook.getSheet("Sổ theo dõi 1 cửa CQ (nhập SL) ");

		while (true) {
			try {
				if (sheet.getRow(currentNumber).getCell(0).getRawValue() == null) {
					break;
				} else if (sheet.getRow(currentNumber).getCell(0).getStringCellValue().compareTo(form.getMaSo()) == 0) {
					this.checkRepetition = 1;
					break;
				}
			} catch (NullPointerException e) {
				break;
			}
			currentNumber++;
		}

		if (this.checkRepetition == 0) {
			row = sheet.createRow(currentNumber);
			cell = row.createCell(0, CellType.STRING);
			cell.setCellValue(form.getMaSo());
			cell = row.createCell(1, CellType.STRING);
			cell.setCellValue(form.getNguonKinhPhi());
			cell = row.createCell(2, CellType.STRING);
			cell.setCellValue(form.getKhoaPhong());
			cell = row.createCell(3, CellType.STRING);
			cell.setCellValue(form.getSoHoatDong());
			cell = row.createCell(4, CellType.STRING);
			cell.setCellValue(form.getNoiDung());
			cell = row.createCell(5, CellType.NUMERIC);
			cell.setCellValue(form.getSoTien());

			inputStream.close();

			// Ghi file
			FileOutputStream out = new FileOutputStream(file);
			try {
				workbook.write(out);
				so1cuathongbao.setText(form.getMaSo() + " đã được thêm vào sổ 1 cửa!");
				out.close();
			} catch (IOException e) {
				e.printStackTrace();
			}

		} else {
			JOptionPane.showMessageDialog(null, "Mã số " + form.getMaSo() + " đã bị trùng!");
		}
		currentNumber = 0;
	}
}
