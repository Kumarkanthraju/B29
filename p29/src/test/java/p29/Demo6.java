package p29;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

//write data to excel
public class Demo6 {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		Workbook wb = WorkbookFactory.create(new FileInputStream("./data/book1.xlsx"));
		wb.getSheet("Sheet1").getRow(0).getCell(0).setCellValue("Bhanu");
		wb.getSheet("Sheet1").getRow(0).createCell(3).setCellValue("Bhanu");
		wb.getSheet("Sheet1").createRow(3).createCell(3).setCellValue("Bhanu");
		wb.write(new FileOutputStream("./data/book2.xlsx"));
		wb.close();
	}
}
