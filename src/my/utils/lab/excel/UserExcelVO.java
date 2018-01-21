package my.utils.lab.excel;

public class UserExcelVO {

	@ExcelColumn(name = "Username")
	private String username;

	@ExcelColumn(name = "Email Address")
	private String email;

	@ExcelColumn(name = "Age")
	private int age;

	@ExcelColumn(name = "Date of Birth")
	private String dob;

	public String getUsername() {
		return username;
	}

	public void setUsername(String username) {
		this.username = username;
	}

	public String getEmail() {
		return email;
	}

	public void setEmail(String email) {
		this.email = email;
	}

	public int getAge() {
		return age;
	}

	public void setAge(int age) {
		this.age = age;
	}

	public String getDob() {
		return dob;
	}

	public void setDob(String dob) {
		this.dob = dob;
	}
}
