package fe.form.excel;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JTextField;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.JTextPane;
import java.awt.Font;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import javax.swing.JOptionPane;
import javax.swing.JLabel;

@SuppressWarnings("serial")
public class GUI extends JFrame {
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
	private JTextPane so1cuathongbao;
	private JTextPane lien1thongbao;
	private JTextPane lien2thongbao;
	private JPanel contentPanel;
	private JFrame frame;
	private JLabel lblChaBit;
	private File testExist = new File("./lock.file");
	public FileInputStream inputStream2;
	public XSSFWorkbook workbookform;
	public XSSFSheet sheetform;

	public static void main(String[] args) throws IOException {
		new GUI();
	}

	public void createTextpaneAndTextfield() {
		contentPanel = new JPanel();

		textPane = new JTextPane();
		textPane.setText("NHẬT  KÝ  LƯU  CHUYỂN  CHỨNG  TỪ");
		textPane.setFont(new Font("Times New Roman", Font.BOLD, 14));
		textPane.setBounds(204, 11, 258, 20);
		contentPanel.add(textPane);

		txtpnChaTpNo = new JTextPane();
		txtpnChaTpNo.setText("Chưa tệp nào được chọn");
		txtpnChaTpNo.setFont(new Font("Times New Roman", Font.PLAIN, 12));
		txtpnChaTpNo.setBounds(122, 42, 288, 25);
		contentPanel.add(txtpnChaTpNo);

		textPane_2 = new JTextPane();
		textPane_2.setText("Mã số:");
		textPane_2.setFont(new Font("Times New Roman", Font.BOLD, 12));
		textPane_2.setBounds(484, 42, 41, 20);
		contentPanel.add(textPane_2);

		textField = new JTextField();
		textField.setColumns(10);
		textField.setBounds(526, 39, 117, 26);
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
		
		so1cuathongbao = new JTextPane();
		so1cuathongbao.setFont(new Font("Times New Roman", Font.PLAIN, 16));
		so1cuathongbao.setBounds(35, 300, 318, 35);
		contentPanel.add(so1cuathongbao);
		
		lien1thongbao = new JTextPane();
		lien1thongbao.setFont(new Font("Times New Roman", Font.PLAIN, 16));
		lien1thongbao.setBounds(35, 346, 330, 35);
		contentPanel.add(lien1thongbao);
		
		lien2thongbao = new JTextPane();
		lien2thongbao.setFont(new Font("Times New Roman", Font.PLAIN, 16));
		lien2thongbao.setBounds(34, 389, 331, 36);
		contentPanel.add(lien2thongbao);
		
		JLabel lblNewLabel = new JLabel("Mã số chứng từ mới nhất: ");
		lblNewLabel.setBounds(383, 313, 169, 26);
		contentPanel.add(lblNewLabel);
		
		lblChaBit = new JLabel("Chưa biết");
		lblChaBit.setBounds(562, 313, 84, 26);
		contentPanel.add(lblChaBit);
	}

	public void createFrame() {
		this.frame = new JFrame();
		this.contentPanel.setLayout(null);
		this.frame.setContentPane(this.contentPanel);
		this.frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		this.frame.getContentPane().setLayout(null);
		this.frame.setBounds(350, 200, 700, 487);
		this.frame.setVisible(true);
		this.frame.addWindowListener(new java.awt.event.WindowAdapter() {
			@Override
			public void windowClosing(java.awt.event.WindowEvent windowEvent) {
				if (JOptionPane.showConfirmDialog(frame, "Bạn có muốn đóng phần mềm này?", "Form With Excel",
						JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE) == JOptionPane.YES_OPTION) {
					testExist.delete();
					System.exit(0);
				}
			}
		});
		this.frame.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
	}

	public void createButton(final So1Cua so1cua, final ChungTu chungtu, final Form form) {
		JButton btnClickMe = new JButton("Vào sổ 1 cửa");
		btnClickMe.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				lien1thongbao.setText("");
				lien2thongbao.setText("");
				so1cuathongbao.setText("");
				if (textField.getText().compareTo("") == 0 || textField_1.getText().compareTo("") == 0
						|| textField_2.getText().compareTo("") == 0 || textField_3.getText().compareTo("") == 0
						|| textField_4.getText().compareTo("") == 0 || textField_5.getText().compareTo("") == 0) {
					JOptionPane.showMessageDialog(null, "Mọi ô trống phải được điền đầy đủ!");
				} else {
					try {
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
								so1cua.updateSo1Cua(form,so1cuathongbao);
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
									chungtu.updateChungTu1(form,lien1thongbao);
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
									chungtu.updateChungTu2(form,lien2thongbao);
								} catch (IOException e1) {
									e1.printStackTrace();
								}
							} else {
								so1cua.setCheckRepe(0);
							}
						}
					} catch (NumberFormatException e) {
						JOptionPane.showMessageDialog(null, "Ô \"Số tiền\" chỉ được nhập số!");
					}
					
				}
			}
		});
		btnClickMe.setBounds(464, 376, 179, 39);
		contentPanel.add(btnClickMe);

		JButton btnChnTp = new JButton("Chọn tệp");
		btnChnTp.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				so1cua.chooseFile(frame,lblChaBit);
				if (so1cua.getFile() != null) {
					txtpnChaTpNo.setText(so1cua.getFile().getName());
					lien1thongbao.setText("");
					lien2thongbao.setText("");
					so1cuathongbao.setText("");
				}
			}
		});
		btnChnTp.setBounds(23, 39, 89, 23);
		contentPanel.add(btnChnTp);
	
	}

	public GUI() throws IOException {		
		if (testExist.exists()) {
			JOptionPane.showMessageDialog(null, "Bạn chỉ được phép mở 1 phần mềm cùng 1 lúc!");
		} else {
			testExist.createNewFile();
			Form form = new Form();
			ChungTu chungtu = new ChungTu();
			So1Cua so1cua = new So1Cua();

			createTextpaneAndTextfield();
			createFrame();
			createButton(so1cua, chungtu, form);
		}
//		testExist.createNewFile();
//		Form form = new Form();
//		ChungTu chungtu = new ChungTu();
//		So1Cua so1cua = new So1Cua();
//
//		createTextpaneAndTextfield();
//		createFrame();
//		createButton(so1cua, chungtu, form);
	}
}
