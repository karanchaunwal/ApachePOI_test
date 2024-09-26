package excelOperations;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel_1 {

	public static void main(String args[]) throws IOException {
		
//		XSSFWorkbook workbook=new XSSFWorkbook();
//		XSSFSheet sheet=workbook.createSheet("sheet_A");
//	
//		Object data[][] ={ 
//			{"Country","Population"},
//			{"India",140},
//			{"China",125} ,
//			{"USA",100} ,
//			{"Indonesia",80}};
//			
//			System.out.println("ROWS: "+data.length);
//			System.out.println("COLUMNS: "+data[0].length);
//			
////			int row=data.length;
////			int column=data[0].length;
//			
//			for(int r=0;r<data.length;r++) {
//				
//				XSSFRow Row=sheet.createRow(r);
//				
//				for(int c=0;c<data[0].length;c++) {
//					
//					XSSFCell cell=Row.createCell(c);
//					Object value=data[r][c];
//					if(value instanceof String) {
//						cell.setCellValue((String)value);
//					}
//
//					if(value instanceof Integer) {
//						cell.setCellValue((Integer)value);
//					}
//
//					
//										
//				}
//			}
//			
//			String filePath=".\\Excel_files\\demo.xlsx";
//			FileOutputStream outputStreamObj= new  FileOutputStream(filePath);
//			workbook.write(outputStreamObj);
//			
//			outputStreamObj.close();
		
		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("Sheet_1");
		Object data[][]= {
				{"Country","Population"},
				{"India",140},
				{"China",125} ,
				{"USA",100} ,
				{"Indonesia",80}
		};
		
		
			for(int r=0;r<data.length;r++) {
				XSSFRow row = sheet.createRow(r);
				for(int c=0;c<data[0].length;c++) {
					XSSFCell cell = row.createCell(c);
					Object value = data[r][c];
					if(value instanceof String) {
						cell.setCellValue((String)value);
					}
					if(value instanceof Integer) {
						cell.setCellValue((Integer)value);
					}
					if(value instanceof Boolean) {
						cell.setCellValue((Boolean)value);
					}
				}
			}
			
			FileOutputStream outputstreamObj=new FileOutputStream(".\\Excel_files\\demo2.xlsx");
			workbook.write(outputstreamObj);
			workbook.close();
			outputstreamObj.close();
			
		
}
	
}