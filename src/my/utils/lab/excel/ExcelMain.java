package my.utils.lab.excel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExcelMain {
	public static void main(String[] args) {
		List<User> userList = new ArrayList<>();

		for (int i = 1; i <= 10000; i++) {
			User user = new User();

			user.setUsername("ABC" + i);
			user.setEmail("abc" + i + "@abc.com");
			user.setAge(20 + i);
			user.setDob(new Date());

			userList.add(user);
		}

		try (FileOutputStream fileOutputStream = new FileOutputStream("/home/klc/Documents/test_" + new Date().getTime() + ".xlsx")) {
			ExcelUtils<UserExcelVO> excelVOExcelUtils = new ExcelUtils<>(UserExcelVO.class);

			long startTime = System.currentTimeMillis();
			excelVOExcelUtils.exportExcel(convertToViewObject(userList), "My User List", fileOutputStream);
			System.out.println("Excel generated! ------- " + (System.currentTimeMillis() - startTime) / 1000 + " seconds");
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	private static List<UserExcelVO> convertToViewObject(List<User> list) {
		List<UserExcelVO> userExcelVOList = new ArrayList<>();
		SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
		for (User user : list) {
			UserExcelVO userExcelVO = new UserExcelVO();

			userExcelVO.setUsername(user.getUsername());
			userExcelVO.setEmail(user.getEmail());
			userExcelVO.setAge(user.getAge());
			if (user.getDob() != null) {
				userExcelVO.setDob(sdf.format(user.getDob()));
			}

			userExcelVOList.add(userExcelVO);
		}

		return userExcelVOList;
	}
}
