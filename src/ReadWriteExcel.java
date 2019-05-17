import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;

import com.anselm.tools.record.Log;

public class ReadWriteExcel {

	public static ArrayList<String> ReadExcelFile(String strFilePath, String strSheetName, String strAPIName, Log log) {
		Workbook workbook = null;
		FileInputStream fis = null;
		ArrayList<String> aryCellList = new ArrayList<>();

		try {
			log.logger("File Path : " + strFilePath);
			fis = new FileInputStream(strFilePath);
			POIFSFileSystem fs = new POIFSFileSystem(fis);
			workbook = new HSSFWorkbook(fs);
			HSSFSheet sheet = (HSSFSheet) workbook.getSheet(strSheetName);

			for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
				HSSFRow row = sheet.getRow(i);
				if (row != null) {
//					Boolean flagCheckStart = false;

					if (row.getCell(1).toString().toUpperCase().trim().equals(strAPIName.toUpperCase())) {
						for (int j = 2; j < row.getLastCellNum(); j++) {
							if (row.getCell(j) == null)
								continue;

							if (row.getCell(j).toString().toUpperCase().trim().equals("END"))
								break;

//							if (row.getCell(j).toString().toUpperCase().equals("START")) {
//								flagCheckStart = true;
//								continue;
//							}
//
//							if (flagCheckStart) {
							aryCellList.add(row.getCell(j).toString().trim());
//							}
						}
						break;
					}
				}
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (fis != null)
					fis.close();

				if (workbook != null)
					workbook.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

		return aryCellList;
	}
}
