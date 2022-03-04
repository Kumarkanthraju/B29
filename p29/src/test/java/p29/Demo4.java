package p29;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/*How to read data from excel file?
 * 1. required dependency--> poi & poi-ooxml
 * 2. poi- poor obfuscation implementation
 * 3. open the excel file
 * 4. goto sheet 
 * 5. goto cell & get the value
 * 6. close the file
 */
public class Demo4 {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		Workbook wb = WorkbookFactory.create(new FileInputStream("./data/book1.xlsx"));
		String v = wb.getSheet("Sheet1").getRow(0).getCell(0).toString();
		System.out.println(v);
		wb.close();
	}
}
