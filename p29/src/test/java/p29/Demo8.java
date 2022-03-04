package p29;
import java.io.FileInputStream;
import java.util.LinkedHashMap;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Demo8 {

	public static void main(String[] args) throws Exception {
		LinkedHashMap<String, String> map=new LinkedHashMap<String, String>();
		  Workbook wb=WorkbookFactory.create(new FileInputStream("./data/book1.xlsx"));
		  int cc=wb.getSheet("sheet2").getRow(0).getLastCellNum();
		  System.out.println(cc);
		  for(int i=0;i<cc;i++) {
			  String k=wb.getSheet("sheet2").getRow(0).getCell(i).toString();
			  String v=wb.getSheet("sheet2").getRow(2).getCell(i).toString();
			  map.put(k, v);  
		  }

		  wb.close();
		  System.out.println(map);
		  
		  System.out.println(map.get("NewPwd"));
	}

}
