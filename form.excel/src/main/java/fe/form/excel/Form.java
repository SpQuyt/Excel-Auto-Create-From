package fe.form.excel;

public class Form {
	private String nguon_kinh_phi;
	private String ma_so;
	private String so_hoat_dong;
	private String noi_dung;
	private long so_tien;
	private String khoa_phong;

	public void setNguonKinhPhi(String nguon_kinh_phi) {
		this.nguon_kinh_phi = nguon_kinh_phi;
	}

	public void setMaSo(String ma_so) {
		this.ma_so = ma_so;
	}

	public void setSoHoatDong(String so_hoat_dong) {
		this.so_hoat_dong = so_hoat_dong;
	}

	public void setNoiDung(String noi_dung) {
		this.noi_dung = noi_dung;
	}

	public void setSoTien(long so_tien) {
		this.so_tien = so_tien;
	}

	public void setKhoaPhong(String khoa_phong) {
		this.khoa_phong = khoa_phong;
	}

	public String getNguonKinhPhi() {
		return this.nguon_kinh_phi;
	}

	public String getMaSo() {
		return this.ma_so;
	}

	public String getSoHoatDong() {
		return this.so_hoat_dong;
	}

	public String getNoiDung() {
		return this.noi_dung;
	}

	public long getSoTien() {
		return this.so_tien;
	}

	public String getKhoaPhong() {
		return this.khoa_phong;
	}
}
