package excelAndGmail;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelCreate {

	public void uploadExcel(Map<String, String> addressSet, List<String> mailList) throws IOException {

		Workbook workbook = null;

		//エクセル内の更新したい欄に記載されているメールアドレスを格納する変数
		String mailAddress = null;

		//エクセル内の値を取得する際に使用する変数
		DataFormatter dataformatter = new DataFormatter();

		File file = new File(Constants.PATH + "TestFile.xlsx");

		ZipSecureFile.setMinInflateRatio(0.001);

		try {
			workbook = WorkbookFactory.create(file);
		}catch(EncryptedDocumentException e) {
			e.printStackTrace();
		}catch(IOException e) {
			e.printStackTrace();
		}

		//修正したセルがわかるように赤縞々で塗りつぶすための設定
		CellStyle styleR = workbook.createCellStyle();
		styleR.setFillBackgroundColor(IndexedColors.RED.getIndex());
		styleR.setFillPattern(FillPatternType.THICK_BACKWARD_DIAG);

		Sheet sheet = workbook.getSheet("Sheet1");

		int lastRow = sheet.getLastRowNum();
		double cellCount = 0;

		FileOutputStream out = null;

		//エクセルデータの最終行まで該当セルの値獲得（メールアドレス値と値変更をするセル）
		for(int i = 0; i <= lastRow; i++) {
			Row row = sheet.getRow(i);

			//アドレスを入力しているセルと値を変更するセルを取得
			Cell addressCell = row.getCell(2);
			Cell targetCell = row.getCell(3);

			//取得したメール送信先のアドレスと同じアドレスがエクセル内にあるかどうかを確認
			if(!addressCell.equals(null)) {
				for(Map.Entry<String, String> entry : addressSet.entrySet()) {
					if(entry.getValue().contains(dataformatter.formatCellValue(addressCell))) {
						mailAddress = entry.getValue();
						break;
					}
				}

				if(mailAddress.equals(Constants.MEMBER)) {
					targetCell.setCellStyle(styleR);
					cellCount = targetCell.getNumericCellValue();
					cellCount++;
					targetCell.setCellValue(cellCount);
				}
			}
			//ループを正常に行うために初期化
			cellCount = 0;
			mailAddress = null;
		}

		//エクセルへの書き込み処理
		try {
			out = new FileOutputStream(Constants.PATH + "TestFileOver.xlsx");
			workbook.write(out);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		} finally{
			out.close();
			workbook.close();
		}
	}
}
