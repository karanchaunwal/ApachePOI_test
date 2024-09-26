package excelOperations;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.Assert;

public class ReadingExcel {

	public static void main(String[] args) throws IOException {
		
//		String excelFilePath=".\\Excel_files\\sample_file.xlsx";
//		
//		FileInputStream inputStreamObj=new FileInputStream(excelFilePath);
//		/*
//				Classes in POI
//				XSSFWorkbook
//				XSSFSheet
//				XSSFRow
//				XSSFCell
//				*/
//		XSSFWorkbook workbook = new XSSFWorkbook(inputStreamObj);
//		
//		
//			XSSFSheet sheet1=workbook.getSheetAt(0); //it will return sheet obj (this func is in XSSFWorkbook class)
//			//or  workbook.getSheet("Sheetname");
//				
//			//****reading sheet from excel
//			//using for loop
//			int LastRowNo=sheet1.getLastRowNum();  //it will return int (this func is in XSSFSheet class)
//			System.out.println("Last Row index: "+LastRowNo);
//			XSSFRow rowObj=sheet1.getRow(1); //it will return Row Obj (this func is in XSSFSheet class)
//			
//			int NoOfColumn=rowObj.getLastCellNum(); //it will return int (this func is in XSSFRow class)
//			System.out.print("Total Columns: "+NoOfColumn);
//			DataFormatter formatter = new DataFormatter();
//			
//			for(int r=0;r<=LastRowNo;r++) {
//				System.out.println("\n");
//				XSSFRow row=sheet1.getRow(r);
//				
//				for(int c=0;c<NoOfColumn;c++) {
//					
//						XSSFCell cellObj=row.getCell(c); //it will return Cell Obj (this func is in XSSFRow class)	
//						//System.out.println("outside if");
//						if (cellObj != null) {
//							
//						//	System.out.println("inside if");
////						cellObj.getCellType(); //(this func is in XSSFCell class)
////				        if(cellObj.getCellType()==CellType.STRING){
////				        	System.out.print(cellObj.getStringCellValue()); //(this func is in XSSFCell class)
////				        	System.out.print(" | ");
////				         }
////				        if(cellObj.getCellType()==CellType.NUMERIC){
////				        	System.out.print(cellObj.getNumericCellValue());
////				        	System.out.print(" | ");
////			   	         }
////				        if(cellObj.getCellType()==CellType.BOOLEAN){
////				        	System.out.print(cellObj.getBooleanCellValue());
////				        	System.out.print(" | ");
////			   	         }
////						if(cellObj.getCellType()==CellType.FORMULA) {
////							System.out.print(cellObj.getNumericCellValue());
////						}
//						String cellValue = formatter.formatCellValue(cellObj);
//						System.out.print(cellValue);
//						if(c<NoOfColumn-1)
//				        System.out.print(" | ");
//				        
//						}
//				        
//				}
//				
//
//		        workbook.close();
//		        inputStreamObj.close();
//			}
		
		
		FileInputStream inputStreamObj = new FileInputStream(".\\Excel_files\\sample_file.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputStreamObj);
		
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		int lastRowno = sheet.getLastRowNum(); //it returns last rows index
		XSSFRow rowObj = sheet.getRow(0);//any row no.
		short lastColumn = rowObj.getLastCellNum(); //it returns last cells number
		
		DataFormatter formatter=new DataFormatter();
		
		for(int r=0;r<=lastRowno;r++) {
			System.out.print("\n");
			XSSFRow row = sheet.getRow(r);
			
			for(int c=0;c<lastColumn;c++) {
				XSSFCell cell = row.getCell(c);
				
				String value = formatter.formatCellValue(cell);
				System.out.print(value);
				if(c<lastColumn-1)
			        System.out.print(" | ");
			        
					
			}
		}
		
	
		
		}
		
		
	
	
}
