package writeexcel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Task8 {

	public static void main(String[] args) throws IOException
	{
         XSSFWorkbook book = new XSSFWorkbook();
         XSSFSheet sheet = book.createSheet("Sheet1");
         Object[][] data = {
        		 {"Name", "Age", "Email" },
        		 {"John Doe", 30, "john@test.com"},
        		 {"Jane Doe", 28,"john@test.com"},
        		 {"Bob Smith", 35,"jacky@example.com"},
        		 {"Swapnil", 37,"swapnil@example.com"}
	};
        int rowcount = 0 ;
        for (Object[] row1 : data )
        {
        	XSSFRow row = sheet.createRow(rowcount++);
        	int columncount = 0;
        	for(Object col : row1)
        	{
        		XSSFCell cell = row.createCell(columncount++);
        		
        		if(col instanceof String)
        		{ cell.setCellValue((Integer)col);}
        		else if (col instanceof Integer)
        		{cell.setCellValue((Integer)col);}
        	}
        }
	}{
	
	try {
		FileOutputStream output = new FileOutputStream("C:\\Users\\hey\\eclipse-workspace\\task8\\Employee details.xlsx");
	book.write(output);
	}
	catch (Exception e)
	{
		e.printStackTrace();
	}
	   book.close();
	
	}
}
