package testapp1;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PoiDemo {

	public static void main(String[] args) {
		
		createworkbook("employees","records");
		readexcel("employees","records");
		appendrow("employees","records");

	}
	
	public static void appendrow(String wb, String ws, String id, String name, String department) {
		
		//check if exists
		File file = new File(wb + ".xlsx");
		if(file.exists()) {
			
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheet(ws);
			int rowlastnum = sheet.getLastRowNum();
			Row newrow = sheet.createRow(rowlastnum+1);
			
			Cell cell1 = newrow.createCell(0);
			cell1.setCellValue(id);
			//newrow.createCell(0).setCellValue(id);
			
			Cell cell2 = newrow.createCell(0);
			cell1.setCellValue(name);
			
			Cell cell3 = newrow.createCell(0);
			cell1.setCellValue(department);
			
			//write to file
			FileOutputStream out = new FileOutputStream(file);
			workbook.write(out);
			System.out.println("new row added");
			out.close();
			
		}
		
	}
	
	//read xlsx
	public static void readworkbook(String workbookname, String worksheetname) {
		
		try {
			File file = new File("employees.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheet("records");
//			XSSFSheet sheet = workbook.getSheet(0);
			
			//loop over rows in sheet
			Iterator<Row> rowiterator = sheet.rowIterator();
			while(rowiterator.hasNext()) {
				Row row = rowiterator.next();
				
				//loop over columns in each row
				Iterator<Cell> celliterator = row.cellIterator();
				while(celliterator.hasNext()) {
					Cell cell = celliterator.next();
					System.out.println(cell.getStringCellValue());
				}//end column loop
			}//end row lopp
			System.out.println("---end---");
			
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
	
	
	public static void createworkbook() {
		//write to xlsx
		//create instance of workbook
		XSSFWorkbook workbook = new XSSFWorkbook(); //HSSFWorkbook sa luma
		XSSFSheet sheet = workbook.createSheet("Employees");
		
		//creae data
		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		data.put("1", new Object[] {"id","name","department"});
		data.put("2", new Object[] {"1","karen","qa"});
		data.put("3", new Object[] {"2","mark","dev"});
		data.put("4", new Object[] {"3","pam","admin"});
				
		Set<String> keyset = data.keySet();
			
		int rownam = 0;
				
		//loop each keyset
		for(String key:keyset) {
					
			Row row = sheet.createRow(rownam+=1);
			Object[] obj = data.get(key);
					
			//loop each column in each row
			int cellnum = 0;
			for (Object o:obj) {
				Cell cell = row.createCell(cellnum+1);
				cell.setCellValue(o.toString());
			}//end of column loop
				}//end of row loop
				
			//write file in filesystem
			try {
			//File file = new File("file.xlsx")
			FileOutputStream out = new FileOutputStream(new File("employees.xlsx"));
			workbook.write(out);
			out.close();
			System.out.println("write xlsx ok");
			} catch (Exception e) {
				System.out.println(e);
			}
	}
	
	public static void createworkbook(String workbookname,String worksheetname) {
		//write to xlsx
		//create instance of workbook
		XSSFWorkbook workbook = new XSSFWorkbook(); //HSSFWorkbook sa luma
		XSSFSheet sheet = workbook.createSheet(worksheetname);
		
		//creae data
		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		data.put("1", new Object[] {"id","name","department"});
		data.put("2", new Object[] {"1","karen","qa"});
		data.put("3", new Object[] {"2","mark","dev"});
		data.put("4", new Object[] {"3","pam","admin"});
			
		Set<String> keyset = data.keySet();
				
		int rownam = 0;
				
		//loop each keyset
		for(String key:keyset) {
					
			Row row = sheet.createRow(rownam+=1);
			Object[] obj = data.get(key);
				
			//loop each column in each row
			int cellnum = 0;
			for (Object o:obj) {
				Cell cell = row.createCell(cellnum+1);
				cell.setCellValue(o.toString());
			}//end of column loop
		}//end of row loop
				
			//write file in filesystem
			try {
				//File file = new File("file.xlsx")
				FileOutputStream out = new FileOutputStream(new File(workbookname + ".xlsx"));
				workbook.write(out);
				out.close();
				System.out.println("write xlsx ok");
			} catch (Exception e) {
					System.out.println(e);
		}
	}
}
