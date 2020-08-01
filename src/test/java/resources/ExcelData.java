package resources;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelData {
	
	
	public static ArrayList<String> getData(String testcasename, String sheetname) throws IOException {
		
		ArrayList<String> list = new ArrayList<String>();
		FileInputStream fis = new FileInputStream("//Users//saurav.singh//Documents//CucumberBDDFramwork//APIFramework//data.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int sheetcount = workbook.getNumberOfSheets();
		for(int i=0; i<sheetcount; i++) {
			if(workbook.getSheetName(i).equalsIgnoreCase(sheetname)) {
			XSSFSheet sheet = workbook.getSheetAt(i);
			
			Iterator<Row> row = sheet.iterator();       // sheet is collection of rows
			Row firstRow = row.next();
			
			Iterator<Cell> cell = firstRow.cellIterator();   // row is collection of cells
			
			int k=0;
			int coloumn = 0;
            while(cell.hasNext()) {
            	
            	Cell value = cell.next();
            	if(value.getStringCellValue().equalsIgnoreCase("TestCases")) 
            	{
            		
            		coloumn = k;
            		
            	}
            	
            	k++;
            } 
            
            System.out.print(coloumn);
            
            while(row.hasNext()) 
            {
              Row r = row.next();
              if(r.getCell(coloumn).getStringCellValue().equalsIgnoreCase(testcasename)) {
            	  
            	  Iterator<Cell> cv = r.iterator();
            	  while(cv.hasNext()) 
            	  {
            		  Cell c = cv.next();
            		  if(c.getCellTypeEnum() == CellType.STRING) {
            		  
            			  list.add(c.getStringCellValue());
            		  
            		  }
            		  else
            		  {
            			  list.add(NumberToTextConverter.toText(c.getNumericCellValue()));
            			  
            		  }
            	  }
            		  
            		  
              }
            	
            }
			}
			
		}
		
		return list;
		

	}

}
