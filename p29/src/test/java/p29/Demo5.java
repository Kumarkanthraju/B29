package p29;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/*How to read data from excel file?
 * 1. required dependency--> poi & poi-ooxml
 * 2. poi- poor obfuscation implementation
 * 3. open the excel file
 * 4. goto sheet 
 * 5. goto cell & get the value
 * 6. close the file
 *how to get cell value?
 *Method overloading example
 *driver.switchTo.frame()
 *asser.asserEqual()
 *wb.getSheet().getRow().getCell().SetCellValue()
 *
 */
public class Demo5 {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		Workbook wb = WorkbookFactory.create(new FileInputStream("./data/book1.xlsx"));
		CellType type1 = wb.getSheet("Sheet2").getRow(0).getCell(0).getCellType();
		System.out.println(type1);

		String A1 = wb.getSheet("Sheet2").getRow(0).getCell(0).getStringCellValue();
		System.out.println(A1);

		String A1a = wb.getSheet("Sheet2").getRow(0).getCell(0).toString();
		System.out.println(A1a);

		CellType type2 = wb.getSheet("Sheet2").getRow(0).getCell(1).getCellType();
		System.out.println(type2);

		double A2 = wb.getSheet("Sheet2").getRow(0).getCell(1).getNumericCellValue();
		System.out.println(A2);

		String A2a = wb.getSheet("Sheet2").getRow(0).getCell(1).toString();
		System.out.println(A2a);

		wb.close();
	}
}
