package fileReader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLReading {
	
	public static Logger appLogs= null;
	
	
	public XLReading() {
		org.apache.log4j.PropertyConfigurator.configure(System.getProperty("user.dir")+ "/log4j.properties");
		appLogs = Logger.getRootLogger();
	}
	

	
	/*
	 * Method: getDataBelowCol
	 * Argument: Workbook Sheet, Column header name 1 and Column header Name 2.
	 * Return the list of String values below the Column header 1 and Column header 2. 
	 * 
	 */
	
	public ArrayList getDataBelowCol(Sheet sheet, String requiredColHeader1, String requiredColHeader2, String runMode){
		int executableSheetRow=0, executableSheetCol=0, cadenceRow=0, cadenceCol=0;
		int totalRows = sheet.getPhysicalNumberOfRows();
		ArrayList sheets = new ArrayList();
		
		for (Row row: sheet){
			for (Cell cell: row){
				cell.setCellType(cell.CELL_TYPE_STRING);
				if(cell.getStringCellValue().equalsIgnoreCase(requiredColHeader1)){
					executableSheetRow= cell.getRowIndex()+1;
					executableSheetCol =cell.getColumnIndex();
					appLogs.debug("Execuatable Sheet column is at Row Number "  + executableSheetRow);
					appLogs.debug("Execuatable Sheet column is at Column Numer " + executableSheetCol);
					continue;
					
				}
				if(cell.getStringCellValue().equalsIgnoreCase(requiredColHeader2)){
				     cadenceRow= cell.getRowIndex()+1;
				     cadenceCol =cell.getColumnIndex();
				     appLogs.debug("Run Mode is at Row Number "+cadenceRow);
					 appLogs.debug("Run Mode is at ColumnNumber " +cadenceCol);
					 continue;
					
				}
			}
						
		}
		
		
		for (int j=executableSheetRow; j<=totalRows; j++)	{
	    	Row row2= sheet.getRow(j);
			if(row2!=null){
				Cell cell1 = row2.getCell(executableSheetCol);
				Cell cell2 = row2.getCell(cadenceCol);
				 if( (cell1!=null && cell2!=null)){
					  if(cell2.getStringCellValue().equalsIgnoreCase(runMode))
					  sheets.add(cell1.getStringCellValue()); 
				 }	
				 else{
					// appLogs.info("Please provide a sheet name at row " + j);
				 }
			}
	    }
		
		return sheets;		
	}
	
	public void getDataBetweenCols(Sheet sheet, String requiredColHeader1, String requiredColHeader2){
		
	}
	
	
	/*
	 * Method: getValueAfterCell
	 * Argument: WorkBook Sheet, Value whose the string whose adjacent (to the right) cell value we want to retrieve.
	 * Find a cell with a particular value (passed as argument) and find the value of the cell next to it. 
	 * return type: String value
	 */
	public String getValueAfterCell(Sheet sheet, String previousCellValue){
		
		String cellValue = "";
		
		for(Row row: sheet){   //For loop to get Browser value starts here
			 int browserCounter= 0;
			 for(Cell cell :row){
				 cell.setCellType(Cell.CELL_TYPE_STRING);
				 if(cell.getStringCellValue().equalsIgnoreCase(previousCellValue)){
					 browserCounter++;	
					 continue;
				 }
				 if(browserCounter==1){
					 cellValue = cell.getStringCellValue();
					 break;
				 }				 				 
				//add more if else for other variables.				 
			 }			 		 		 
		 }//For Loop to Get Browser value ends here.
		
		return cellValue;
		
	}
	
	
	
	/*
	 * Method: verifySheetPresence
	 * Arguments: Workbook object, Sheet name
	 * verifies if a particular sheet is present in the file or not.
	 * returns true is sheet is found, else returns false.
	 * 
	 */
	public boolean verifySheetPresence(Workbook xlReader, String sheetName){
		
		boolean sheetFound= false;
		   for( int j=0; j<xlReader.getNumberOfSheets(); j++){
			   		  
			   if(sheetName.trim().equals(xlReader.getSheetName(j))){
				   sheetFound = true;
				  appLogs.debug(sheetName + " is present in the file.");
			   }		  
		   }
		   
		 return sheetFound;
		
	}
	
	
	/*
	 * Method: columnHeaderIndex
	 * Arguments: Sheet object, column Header name whose index is to be found.
	 * Finds the column header index in a sheet.
	 * returns the column header index if found, else returns 0
	 * 
	 */
	
	public int columnHeaderIndex(Sheet sheet , String columnHeader){
		
		for(Cell cell : sheet.getRow(0)){   // Loop to find the column index position of all the mandatory Column headers starts here
	    	 cell.setCellType(Cell.CELL_TYPE_STRING);
	    	   //appLogs.info("header Name:" + cell.getStringCellValue());
	    	if(cell.getStringCellValue().trim().equalsIgnoreCase(columnHeader)){
	    		return cell.getColumnIndex();
	    		//appLogs.info("Name column is at index " + nameIndex);
	    	} 
	    			
	}// Loop to find the column index position of all the mandatory Column headers starts here
		 return 0;
	}
	
	
	
	/*
	 * Method: POIObjectCreator
	 * This method is used to support xls and xlsx sheet format
	 * It returns the Workbook object based on the type of extension of the excel file.
	 * 
	 */
	public Workbook POIObjectCreator(File file, String fileName, FileInputStream inputstream){
		Workbook xlReader= null;
		try {
			FileInputStream inputStream = new FileInputStream(file);
			//String fileExtension = fileName.substring(fileName.indexOf("."));
			String fileExtension = file.getName().substring(file.getName().indexOf("."));
			if(fileExtension.equalsIgnoreCase(".xls")){
				try {
					xlReader = new HSSFWorkbook(inputStream);
					appLogs.debug("File " + fileName + " has .xls extension ");
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
			else if(fileExtension.equalsIgnoreCase(".xlsx")){
				try {
					xlReader = new XSSFWorkbook(inputStream);
					appLogs.debug("File " + fileName + " has .xlsx extension ");
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return xlReader;
		
		
		
	}
	
	
	/* 
	 * Method: getCellData
	 * its generic method to get data from a aparticular cell in the excel sheet
	 * It takes the Workbook object, SheetName, rowIndex and ColumnIndex of a particular sheet 
	 * Returns the String value at that particular row and column index that are passed as parameters.
	 * 
	 */
	
	public String getCellData(Workbook xlreader, String sheetName,  int rowIndex, int columnIndex){
		
		//appLogs.info("sheet "  +sheetName);
		Sheet sheet1 = xlreader.getSheet(sheetName);
		for (Row row : sheet1){
			for (Cell cell : row){
				cell.setCellType(Cell.CELL_TYPE_STRING);
				if(cell!=null){
					if((row.getRowNum()== rowIndex) && (cell.getColumnIndex()== columnIndex)){
						if(cell.getStringCellValue().isEmpty())
							return "N/A";					
						else
		 				return cell.getStringCellValue().trim();					
					}
				}			
			}
		}
		
		return null;
		
	}
	
	
	/*
	 * Method: getPhysicalRows
	 * Argument: Workbook object, SheetName, index number of a column header, the type of sheet we are looking into (main or common)
	 * Returns the total number of rows present in the sheet. 
	 * return type: integer
	 */
	
	public int getPhysicalRows(Workbook xlreader, String sheetName , int actionIndex , String sheetType){
		int counter= 0;
		Sheet sheet1 = xlreader.getSheet(sheetName);
		int loopTill= sheet1.getPhysicalNumberOfRows()+10;
		
		for (int i=1; i<=sheet1.getPhysicalNumberOfRows();i++){
			Row row = sheet1.getRow(i);
			Cell cell = null;
			if(row!=null){
				cell = row.getCell(actionIndex);	}
			
			if(cell!=null){
				if( cell.getStringCellValue().equals("quit") && sheetType.equals("mainSheet") ){
								return cell.getRowIndex();
				}
				if( cell.getCellType() == Cell.CELL_TYPE_STRING && cell.getStringCellValue().isEmpty() && sheetType.equals("commonSheet") ){
					return cell.getRowIndex();
				}
						
				
			}				
			
		}
		
		return counter;
	}
	
	
	
}

