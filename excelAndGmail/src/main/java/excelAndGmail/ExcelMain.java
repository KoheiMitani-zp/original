package excelAndGmail;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.List;

public class ExcelMain {

	//エクセル効率化業務メインクラス
	public static void main(String[] args) throws IOException, GeneralSecurityException {

		ExcelControll excelControll = new ExcelControll();

		GmailControll gmailControll = new GmailControll();

		ExcelCreate excelCreate = new ExcelCreate();

		//今回の該当メールアドレスを事前に取得
		List<String> mailList = excelControll.getFile();

		excelCreate.uploadExcel(gmailControll.gmailConnect(mailList), mailList);

	}

}