package mavenProjectTest.mavenProj;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadingData {

	public static void main(String[] args) throws IOException {
		
		ArrayList<String> a = new ArrayList<String>();

		FileInputStream fis = new FileInputStream("E:\\Selenium_Workspace\\SeleniumBasics\\Testdatasheet.xlsx");

		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		int sheets = workbook.getNumberOfSheets();

		for (int i = 0; i < sheets; i++)

		{
			if (workbook.getSheetName(i).equalsIgnoreCase("Testdata"))

			{

				XSSFSheet sheet = workbook.getSheetAt(i);
				
				Iterator<Row> row = sheet.iterator();
				
				Row firstrow = row.next();
				
				Iterator<Cell> ce = firstrow.cellIterator();
				
				int k =0;
				
				int column = 0;
				
				while(ce.hasNext())
					
				{
					Cell value = ce.next();
					
					if(value.getStringCellValue().equalsIgnoreCase("Testcase"))
					
					{
						
						column=k;
						
						
					}
					
					k++;
				}
				
				//System.out.println(column);
				
				
				while(row.hasNext())
					
				{
					
					Row r = row.next();
					
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase("Purchase"))
						
					{
						Iterator<Cell> re = r.cellIterator();
						
						while(re.hasNext())
							
						{
							
							Cell ce1 = re.next();
							
							 a.add(ce1.getStringCellValue());
							 
							 
						}
						
						System.out.println(a);
						
					}
				}
				
			}
				

		}

	}

}
