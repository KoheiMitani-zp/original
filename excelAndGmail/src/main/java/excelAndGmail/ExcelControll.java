package excelAndGmail;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

//エクセルデータの取得
public class ExcelControll {

	public List<String> getFile() throws IOException{

		List<String> mailList = new ArrayList<String>();

		//create new workbook
		Workbook workbook = null;

		FileOutputStream out = null;

		//filePath
		File file = new File(Constants.PATH + "TestFile.xlsx");

		//change file size capacity
		ZipSecureFile.setMinInflateRatio(0.001);

		try {
			workbook = WorkbookFactory.create(file);
		} catch(EncryptedDocumentException e){
			e.printStackTrace();
		} catch(IOException e) {
			e.printStackTrace();
		}

		Sheet sheet = workbook.getSheet("Sheet1");

		int lastRow = sheet.getLastRowNum();

		int lastCol = 3;

		//データのあるすべての行を確認
		for(int i = 0; i <= lastRow; i++) {

			Row row = sheet.getRow(i);

			Cell nameCell = row.getCell(1);
			Cell addressCell = row.getCell(2);

			for(int n = 0; n <= lastCol; n++) {

				Cell allCell = row.getCell(n);

				if(allCell != null) {
					CellStyle whiteStyel = workbook.createCellStyle();
					//セルの背景を白色に変更
					whiteStyel.setFillBackgroundColor(IndexedColors.WHITE.getIndex());
					whiteStyel.setFillPattern(FillPatternType.NO_FILL);
					allCell.setCellStyle(whiteStyel);
				}else {
					continue;
				}
			}

			if(nameCell != null && addressCell != null) {
				//メールアドレスをコレクションに追加
				String mailAddress = addressCell.getStringCellValue();
				mailList.add(mailAddress);
			}
		}

		try {
			out = new FileOutputStream(Constants.PATH + "TestFileOver.xlsx");
			workbook.write(out);
		}catch(Exception e) {
			e.printStackTrace();
		}finally {
			out.close();
			workbook.close();
		}

		for(String mail : mailList) {
			System.out.println(mail);
		}

		return mailList;
	}
}
