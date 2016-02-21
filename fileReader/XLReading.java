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
	
	
	/**
	 * Class Constructor
	 * 
	 */
	public XLReading() {
		org.apache.log4j.PropertyConfigurator.configure(System.getProperty("user.dir")+ "/log4j.properties");
		appLogs = Logger.getRootLogger();
	}
	
	
	/**
	 * Scan the excel sheet and retrieves the data below two column headers based on the 
	 * provided cadence
	 * 
	 * @param sheet - WorkBook Sheet object for the currently running Sheet
	 * @param requiredColHeader1 - String value of the Column header
	 * @param requiredColHeader2- String value of the column header
	 * @param runMode - Cadence - Daily, Weekly or Release
	 * @return - arrayList of String based on the provided cadence. 
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
	
	
	/**
	 * Scan the excel sheet and retrieves the data form the cell to the right of the provided cell value
	 * 
	 * @param sheet - WorkBook Sheet object for the currently running Sheet
	 * @param previousCellValue -The cell value of the previous cell. 
	 * @return - a String cell value after the specified cell value.
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
	
	
	
	/**
	 * Scan the excel sheet to check if the particular sheet name provided in the config is present in the file
	 * 
	 * @param xlReader: Workbook object of the currently running file. 
	 * @param SheetName - Name of the sheet the presence of which needs to be tested.
	 * @return True if the sheet in present and false otherwise.
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
	
	
	/**
	 * Scan the excel sheet to find the position of a particular column header.
	 * 
	 * @param sheet - WorkBook Sheet object for the currently running Sheet
	 * @param columnHeader - String value of the column header position of which needs to be found.
	 * @return - integer value of the column header position in the excel sheet.
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
	
	
	
	
	/**
	 * Checks the file extension then then creates a Workbook object for that file.
	 * 
	 * @param file - File object of the currently running File
	 * @param FileName - Name of the currently running file.
	 * @param inputStream - inputStream object to interact with the file.
	 * @return - a WorkBook object of the File. 
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
	
	/**
	 * Fetches the data from a particular cell 
	 * 
	 * @param xlreader - WorkBook Sheet object for the currently running Sheet
	 * @param sheetName- Name of the currently Running sheet inside the file.
	 * @param rowIndex- row index of the cell from which the data is to be fetched.
	 * @param columnIndex -column index of the cell from which the data is to be fetched.
	 * @return - returns a String value form the cell. 
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
	
	
	/**
	 * gets the total physical number of rows in a sheet. 
	 * 
	 * @param xlreader - WorkBook Sheet object for the currently running Sheet
	 * @param sheetName- Name of the currently Running sheet inside the file.
	 * @param actionIndex- index value of the 'Action' column header.
	 * @param SheetType - sheet from the Main folder or common folder.
	 * @return - returns integer value of the total number of rows in the column. 
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

