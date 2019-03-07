package fe.form.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextPane;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class So1Cua {
	private int approve = 0;
	private File file;
	private JLabel label;
	public static int currentNumber = 7;
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
	
	public void checkLatestCode() throws IOException {
		Cell cell;
		Row row;
		CellStyle style;
		CellStyle[] stylearray;
		XSSFFont font;

		FileInputStream inputStream = new FileInputStream(this.file);
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet sheet = workbook.getSheet("Sổ theo dõi 1 cửa CQ (nhập SL) ");
		font = workbook.createFont();
		font.setFontName("Times New Roman");
		font.setFontHeight(8);

		while (true) {
			try {
				if (sheet.getRow(currentNumber).getCell(0).getCellTypeEnum() == CellType.BLANK
						|| sheet.getRow(currentNumber).getCell(0).getCellTypeEnum() == CellType._NONE) {
					this.label.setText(sheet.getRow(currentNumber-1).getCell(0).getStringCellValue());
					break;
				}
			} catch (NullPointerException e) {
				break;
			}
			currentNumber++;
		}
		currentNumber = 7;
		inputStream.close();
	}

	public void chooseFile(JFrame frame, JLabel lblChaBit) {
		this.label = lblChaBit;
		JFileChooser chooser = new JFileChooser("./");
		FileNameExtensionFilter filter = new FileNameExtensionFilter("Chỉ được chọn file excel...", "xlsx", "xls");
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
		
		try {
			checkLatestCode();
		} catch (IOException e) {
			e.printStackTrace();
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
		CellStyle[] stylearray;
		XSSFFont font;

		FileInputStream inputStream = new FileInputStream(this.file);
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet sheet = workbook.getSheet("Sổ theo dõi 1 cửa CQ (nhập SL) ");
		font = workbook.createFont();
		font.setFontName("Times New Roman");
		font.setFontHeight(8);

		while (true) {
			try {
				if (sheet.getRow(currentNumber).getCell(0).getCellTypeEnum() == CellType.NUMERIC) {
					sheet.getRow(currentNumber).getCell(0).setCellType(CellType.STRING);
					if (sheet.getRow(currentNumber).getCell(0).getStringCellValue().compareTo(form.getMaSo()) == 0) {
						this.checkRepetition = 1;
						break;
					}
				} else if (sheet.getRow(currentNumber).getCell(0).getCellTypeEnum() == CellType.STRING) {
					if (sheet.getRow(currentNumber).getCell(0).getStringCellValue().compareTo(form.getMaSo()) == 0) {
						this.checkRepetition = 1;
						break;
					}
				} else if (sheet.getRow(currentNumber).getCell(0).getCellTypeEnum() == CellType.BLANK
						|| sheet.getRow(currentNumber).getCell(0).getCellTypeEnum() == CellType._NONE) {
					this.label.setText(sheet.getRow(currentNumber).getCell(0).getStringCellValue());
					break;
				}
			} catch (NullPointerException e) {
				break;
			}
			currentNumber++;
		}

		if (this.checkRepetition == 0) {
			stylearray = new CellStyle[13];
			for (int i = 0; i < 13; i++) {
				stylearray[i] = sheet.getRow(currentNumber - 1).getCell(i).getCellStyle();
			}

			row = sheet.createRow(currentNumber);

			cell = row.createCell(0, CellType.STRING);
			cell.setCellStyle(stylearray[0]);
			cell.setCellValue(form.getMaSo());

			cell = row.createCell(1, CellType.STRING);
			cell.setCellStyle(stylearray[1]);
			cell.setCellValue(form.getNguonKinhPhi());

			cell = row.createCell(2, CellType.STRING);
			cell.setCellStyle(stylearray[2]);
			cell.setCellValue(form.getKhoaPhong());

			cell = row.createCell(3, CellType.STRING);
			cell.setCellStyle(stylearray[3]);
			cell.setCellValue(form.getSoHoatDong());

			cell = row.createCell(4, CellType.STRING);
			cell.setCellStyle(stylearray[4]);
			cell.setCellValue(form.getNoiDung());

			cell = row.createCell(5, CellType.STRING);
			cell.setCellStyle(stylearray[5]);
			cell.setCellValue(form.getSoTien());
			
			cell = row.createCell(6, CellType.STRING);
			cell.setCellStyle(stylearray[6]);
			
			cell = row.createCell(7, CellType.STRING);
			cell.setCellStyle(stylearray[7]);
			
			cell = row.createCell(8, CellType.STRING);
			cell.setCellStyle(stylearray[8]);
			
			cell = row.createCell(9, CellType.STRING);
			cell.setCellStyle(stylearray[9]);
			
			cell = row.createCell(10, CellType.STRING);
			cell.setCellStyle(stylearray[10]);
			
			cell = row.createCell(11, CellType.STRING);
			cell.setCellStyle(stylearray[11]);
			
			cell = row.createCell(12, CellType.STRING);
			cell.setCellStyle(stylearray[12]);

			inputStream.close();

			// Ghi file
			FileOutputStream out = new FileOutputStream(file);
			try {
				workbook.write(out);
				this.label.setText(sheet.getRow(currentNumber).getCell(0).getStringCellValue());
				so1cuathongbao.setText(form.getMaSo() + " đã được thêm vào sổ 1 cửa!");
				out.close();
			} catch (IOException e) {
				e.printStackTrace();
			}

		} else {
			JOptionPane.showMessageDialog(null, "Mã số " + form.getMaSo() + " đã bị trùng!");
		}
		currentNumber = 7;
	}
}
