package p29;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/*How to read data from excel file?
 * 1. required dependency--> poi & poi-ooxml
 * 2. poi- poor obfuscation implementation
 * 3. 
 */
public class Demo3 {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		// open the excel file
		Workbook wb = WorkbookFactory.create(new FileInputStream("./Data/Book1.xlsx"));
		// goto sheet1
		Sheet s = wb.getSheet("sheet1");
		// goto 1st row 1st col, get the value
		String v = s.getRow(0).getCell(0).getStringCellValue();
		System.out.println(v);
		// close the xl file
		wb.close();

	}
}
