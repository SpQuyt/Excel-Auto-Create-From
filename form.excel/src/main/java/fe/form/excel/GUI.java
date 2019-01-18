package fe.form.excel;

import java.awt.EventQueue;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextField;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.JTextPane;
import java.awt.Font;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.awt.print.PageFormat;
import java.awt.print.Printable;
import java.awt.print.PrinterException;
import java.awt.print.PrinterJob;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.awt.Color;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;
import javax.swing.border.LineBorder;
import javax.swing.JLabel;
import javax.swing.JOptionPane;

import java.awt.Canvas;

@SuppressWarnings("serial")
public class GUI extends JFrame {

	JTextField field;
	JButton button;
	private JTextField textField;
	private JTextField textField_1;
	private JTextField textField_2;
	private JTextField textField_3;
	private JTextField textField_4;
	private JTextField textField_5;
	private JTextPane textPane;
	private JTextPane txtpnChaTpNo;
	private JTextPane textPane_2;
	private JTextPane textPane_3;
	private JTextPane textPane_4;
	private JTextPane textPane_5;
	private JTextPane textPane_6;
	private JTextPane textPane_7;
	private JTextPane textPane_8;
	private JPanel contentPanel;
	public FileInputStream inputStream2;
	public XSSFWorkbook workbookform;
	public XSSFSheet sheetform;

	public static void main(String[] args) throws IOException {
		new GUI();
	}

	public void initialize() {
		contentPanel = new JPanel();

		textPane = new JTextPane();
		textPane.setText("NHẬT  KÝ  LƯU  CHUYỂN  CHỨNG  TỪ");
		textPane.setFont(new Font("Times New Roman", Font.BOLD, 14));
		textPane.setBounds(204, 11, 258, 20);
		contentPanel.add(textPane);

		txtpnChaTpNo = new JTextPane();
		txtpnChaTpNo.setText("Chưa tệp nào được chọn");
		txtpnChaTpNo.setFont(new Font("Times New Roman", Font.PLAIN, 12));
		txtpnChaTpNo.setBounds(122, 42, 135, 20);
		contentPanel.add(txtpnChaTpNo);

		textPane_2 = new JTextPane();
		textPane_2.setText("Mã số:");
		textPane_2.setFont(new Font("Times New Roman", Font.BOLD, 12));
		textPane_2.setBounds(484, 45, 41, 20);
		contentPanel.add(textPane_2);

		textField = new JTextField();
		textField.setColumns(10);
		textField.setBounds(526, 33, 117, 32);
		contentPanel.add(textField);

		textField_1 = new JTextField();
		textField_1.setColumns(10);
		textField_1.setBounds(143, 76, 503, 39);
		contentPanel.add(textField_1);

		textPane_3 = new JTextPane();
		textPane_3.setText("Nguồn kinh phí:");
		textPane_3.setFont(new Font("Times New Roman", Font.BOLD, 12));
		textPane_3.setBounds(35, 83, 98, 20);
		contentPanel.add(textPane_3);

		textPane_4 = new JTextPane();
		textPane_4.setText("Số HĐ:");
		textPane_4.setFont(new Font("Times New Roman", Font.BOLD, 12));
		textPane_4.setBounds(35, 165, 52, 20);
		contentPanel.add(textPane_4);

		textField_2 = new JTextField();
		textField_2.setColumns(10);
		textField_2.setBounds(93, 144, 186, 63);
		contentPanel.add(textField_2);

		textPane_5 = new JTextPane();
		textPane_5.setText("Nội dung:");
		textPane_5.setFont(new Font("Times New Roman", Font.BOLD, 12));
		textPane_5.setBounds(291, 165, 59, 20);
		contentPanel.add(textPane_5);

		textField_3 = new JTextField();
		textField_3.setColumns(10);
		textField_3.setBounds(350, 144, 296, 63);
		contentPanel.add(textField_3);

		textPane_6 = new JTextPane();
		textPane_6.setText("Số tiền:");
		textPane_6.setFont(new Font("Times New Roman", Font.BOLD, 12));
		textPane_6.setBounds(35, 250, 52, 20);
		contentPanel.add(textPane_6);

		textField_4 = new JTextField();
		textField_4.setColumns(10);
		textField_4.setBounds(93, 234, 186, 55);
		contentPanel.add(textField_4);

		textPane_7 = new JTextPane();
		textPane_7.setText("Khoa/Phòng:");
		textPane_7.setFont(new Font("Times New Roman", Font.BOLD, 12));
		textPane_7.setBounds(291, 251, 74, 20);
		contentPanel.add(textPane_7);

		textField_5 = new JTextField();
		textField_5.setColumns(10);
		textField_5.setBounds(370, 234, 276, 55);
		contentPanel.add(textField_5);

		textPane_8 = new JTextPane();
		textPane_8.setText("Mẫu số KSBTTPHN.01");
		textPane_8.setFont(new Font("Times New Roman", Font.PLAIN, 9));
		textPane_8.setBounds(548, 11, 98, 20);
		contentPanel.add(textPane_8);
	}

	public GUI() throws IOException {
		final Form form = new Form();
		final ChungTu chungtu = new ChungTu();
		final So1Cua so1cua = new So1Cua();
//		final FileWriter fw;

		JOptionPane.showMessageDialog(null, "Hãy tắt Sổ 1 Cửa trước khi sử dụng phần mềm nhé!");

		this.initialize();

		final JFrame frame = new JFrame();
		contentPanel.setLayout(null);
		frame.setContentPane(contentPanel);
		frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		frame.setBounds(350, 200, 700, 400);
		frame.setVisible(true);

		JButton btnClickMe = new JButton("Vào sổ 1 cửa");
		btnClickMe.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if (textField.getText().compareTo("") == 0 || textField_1.getText().compareTo("") == 0
						|| textField_2.getText().compareTo("") == 0 || textField_3.getText().compareTo("") == 0
						|| textField_4.getText().compareTo("") == 0 || textField_5.getText().compareTo("") == 0) {
					JOptionPane.showMessageDialog(null, "Mọi ô trống phải được điền đầy đủ!");
				} else {
					form.setMaSo(textField.getText());
					form.setNguonKinhPhi(textField_1.getText());
					form.setSoHoatDong(textField_2.getText());
					form.setNoiDung(textField_3.getText());
					form.setSoTien(Long.parseLong(textField_4.getText()));
					form.setKhoaPhong(textField_5.getText());

					// sổ 1 cửa
					if (so1cua.getFile() == null) {
						JOptionPane.showMessageDialog(null, "Chưa tệp nào được chọn!");
					} else {
						System.out.println(so1cua.getFile());
						try {
							so1cua.updateSo1Cua(form);
						} catch (IOException e) {
							e.printStackTrace();
						}

						if (so1cua.getCheckRepe() == 0) {
							// nhật ký lưu chuyển chứng từ liên 1
							try {
								chungtu.createChungTu1(form.getMaSo());
							} catch (IOException e1) {
								e1.printStackTrace();
							}
							try {
								chungtu.updateChungTu1(form);
							} catch (IOException e1) {
								e1.printStackTrace();
							}
							
							// nhật ký lưu chuyển chứng từ liên 2
							try {
								chungtu.createChungTu2(form.getMaSo());
							} catch (IOException e1) {
								e1.printStackTrace();
							}
							try {
								chungtu.updateChungTu2(form);
							} catch (IOException e1) {
								e1.printStackTrace();
							}
						} else {
							so1cua.setCheckRepe(0);
						}
					}
				}
			}
		});
		btnClickMe.setBounds(456, 312, 179, 39);
		contentPanel.add(btnClickMe);

		JButton btnChnTp = new JButton("Chọn tệp");
		btnChnTp.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				so1cua.chooseFile(frame);
				if (so1cua.getFile() != null) {
					txtpnChaTpNo.setText(so1cua.getFile().getName());
				}
			}
		});
		btnChnTp.setBounds(23, 39, 89, 23);
		contentPanel.add(btnChnTp);
	}
}
