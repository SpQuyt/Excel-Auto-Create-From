package fe.form.excel;

import java.io.*;

import javax.swing.JOptionPane;
import javax.swing.JTextPane;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ChungTu {
	private File fileForm;

	public File getFileForm() {
		return this.fileForm;
	}

	public void setFileForm(File fileForm) {
		this.fileForm = fileForm;
	}

	private static void copyFileUsingStream(File source, File dest) throws IOException {
		InputStream is = null;
		OutputStream os = null;
		try {
			is = new FileInputStream(source);
			os = new FileOutputStream(dest);
			byte[] buffer = new byte[1024];
			int length;
			while ((length = is.read(buffer)) > 0) {
				os.write(buffer, 0, length);
			}
		} finally {
			is.close();
			os.close();
		}
	}

	public void createChungTu1(String name) throws IOException {
		File source = new File("./sampleForm1.xlsx");
		File dest = new File("./new/" + name + "_Lien1.xlsx");
		long start = System.nanoTime();
		copyFileUsingStream(source, dest);
		this.fileForm = dest;
		System.out.println("Time taken by Stream Copy = " + (System.nanoTime() - start));
	}

	public void updateChungTu1(Form form, JTextPane lien1thongbao) throws IOException {
		Cell cell;
		Row row;
		CellStyle style;
		XSSFFont font;

		FileInputStream inputStream = new FileInputStream(this.fileForm);
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		font = workbook.createFont();
		font.setFontName("Times New Roman");
		font.setFontHeight(10);

		row = sheet.getRow(2);
		cell = row.createCell(8, CellType.STRING);
		style = cell.getCellStyle();
		style.setFont(font);
		style.setVerticalAlignment(VerticalAlignment.TOP);
		cell.setCellValue(form.getMaSo());
		
		row = sheet.getRow(2);
		cell = row.createCell(2, CellType.STRING);
		style = cell.getCellStyle();
		style.setFont(font);
		style.setVerticalAlignment(VerticalAlignment.TOP);
		cell.setCellValue(form.getNguonKinhPhi());
		
		row = sheet.getRow(2);
		cell = row.createCell(8, CellType.STRING);
		style = cell.getCellStyle();
		style.setFont(font);
		style.setVerticalAlignment(VerticalAlignment.TOP);
		cell.setCellValue(form.getMaSo());
		
		row = sheet.getRow(4);
		cell = row.createCell(2, CellType.STRING);
		style = cell.getCellStyle();
		style.setFont(font);
		style.setVerticalAlignment(VerticalAlignment.TOP);
		cell.setCellValue(form.getSoHoatDong());
		
		row = sheet.getRow(4);
		cell = row.createCell(4, CellType.STRING);
		style = cell.getCellStyle();
		style.setFont(font);
		style.setVerticalAlignment(VerticalAlignment.TOP);
		cell.setCellValue(form.getNoiDung());
		
		row = sheet.getRow(6);
		cell = row.createCell(2, CellType.NUMERIC);
		style = cell.getCellStyle();
		style.setFont(font);
		style.setVerticalAlignment(VerticalAlignment.TOP);
		cell.setCellValue(form.getSoTien());
		
		row = sheet.getRow(6);
		cell = row.createCell(4, CellType.STRING);
		style = cell.getCellStyle();
		style.setFont(font);
		style.setVerticalAlignment(VerticalAlignment.TOP);
		cell.setCellValue(form.getKhoaPhong());
		
		inputStream.close();
		
		// Ghi file
		FileOutputStream out = new FileOutputStream(fileForm);
		try {
			workbook.write(out);
			lien1thongbao.setText(form.getMaSo() + " Liên 1 đã được tạo mới!");
			out.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public void createChungTu2(String name) throws IOException {
		File source = new File("./sampleForm2.xlsx");
		File dest = new File("./new/" + name + "_Lien2.xlsx");
		long start = System.nanoTime();
		copyFileUsingStream(source, dest);
		this.fileForm = dest;
		System.out.println("Time taken by Stream Copy = " + (System.nanoTime() - start));
	}

	public void updateChungTu2(Form form, JTextPane lien2thongbao) throws IOException {
		Cell cell;
		Row row;
		CellStyle style;
		XSSFFont font;

		FileInputStream inputStream = new FileInputStream(this.fileForm);
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		font = workbook.createFont();
		font.setFontName("Times New Roman");
		font.setFontHeight(10);

		row = sheet.getRow(2);
		cell = row.createCell(8, CellType.STRING);
		style = cell.getCellStyle();
		style.setFont(font);
		style.setVerticalAlignment(VerticalAlignment.TOP);
		cell.setCellValue(form.getMaSo());
		
		row = sheet.getRow(2);
		cell = row.createCell(2, CellType.STRING);
		style = cell.getCellStyle();
		style.setFont(font);
		style.setVerticalAlignment(VerticalAlignment.TOP);
		cell.setCellValue(form.getNguonKinhPhi());
		
		row = sheet.getRow(2);
		cell = row.createCell(8, CellType.STRING);
		style = cell.getCellStyle();
		style.setFont(font);
		style.setVerticalAlignment(VerticalAlignment.TOP);
		cell.setCellValue(form.getMaSo());
		
		row = sheet.getRow(4);
		cell = row.createCell(2, CellType.STRING);
		style = cell.getCellStyle();
		style.setFont(font);
		style.setVerticalAlignment(VerticalAlignment.TOP);
		cell.setCellValue(form.getSoHoatDong());
		
		row = sheet.getRow(4);
		cell = row.createCell(4, CellType.STRING);
		style = cell.getCellStyle();
		style.setFont(font);
		style.setVerticalAlignment(VerticalAlignment.TOP);
		cell.setCellValue(form.getNoiDung());
		
		row = sheet.getRow(6);
		cell = row.createCell(2, CellType.NUMERIC);
		style = cell.getCellStyle();
		style.setFont(font);
		style.setVerticalAlignment(VerticalAlignment.TOP);
		cell.setCellValue(form.getSoTien());
		
		row = sheet.getRow(6);
		cell = row.createCell(4, CellType.STRING);
		style = cell.getCellStyle();
		style.setFont(font);
		style.setVerticalAlignment(VerticalAlignment.TOP);
		cell.setCellValue(form.getKhoaPhong());
		
		inputStream.close();
		
		// Ghi file
		FileOutputStream out = new FileOutputStream(fileForm);
		try {
			workbook.write(out);
			lien2thongbao.setText(form.getMaSo() + " Liên 2 đã được tạo mới!");
			out.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public static void main(String[] args) throws IOException {
		ChungTu test = new ChungTu();
		test.createChungTu1("huh");
	}
}
