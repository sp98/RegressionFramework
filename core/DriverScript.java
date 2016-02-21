package core;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.testng.annotations.Test;

import reporting.HTMLReport;
import utilities.Utils;
import fileReader.XLReading;



public class DriverScript extends XLReading  {

	
	private String inputDirectory = null;
	private String commonInputDirectory = null;
	private String masterFileName = "MasterSheet";
	private String configSheet = "config";
	private String runMode = "";

	/*Result Set Data Structure*/
	private static LinkedHashMap<String, ArrayList<String>> resultSet = null;
	private ArrayList<String> subResultSet = null;
	
	private ArrayList<String> executableFiles = null;
	private HashMap<String, LinkedHashMap<String, String>> fileVariables = null;
	private ArrayList<String> executableSheets = null;
	private HashMap<String, HashMap<Workbook, ArrayList<String>>> allXlReaderObjs = null;
	private HashMap<Workbook, ArrayList<String>> xlReaderObjs = null;

	private LinkedHashMap<String, ArrayList<ArrayList<String>>> sheetData = null;
	private ArrayList<ArrayList<String>> sheetRowData = null;
	private ArrayList<String> rowData = null;
	private ArrayList<String> commonSheetRowData = null;
	
	/*Hashmap that contains the runtime variables created in the common Sheet */
	private LinkedHashMap<String , HashMap<String, String>> allCommonSheetVariables = null;
	HashMap<String, String> commonSheetVariables= null;

	private HTMLReport htmlReport = null;


	/**
	 * Class Constructor
	 * 
	 */
	public DriverScript() {
		
		/*initializing the working Directories */
		inputDirectory = System.getProperty("user.dir") + "/Input/";
		commonInputDirectory = System.getProperty("user.dir")+ "/Input/Common/";
		
		runMode = "Daily"; // take this as user argument
		
		/*initializing the global data Structures */
		fileVariables = new HashMap<String, LinkedHashMap<String, String>>();
		allXlReaderObjs = new HashMap<String, HashMap<Workbook, ArrayList<String>>>();
		sheetData = new LinkedHashMap<String, ArrayList<ArrayList<String>>>();
		allCommonSheetVariables = new LinkedHashMap<String, HashMap<String, String>>();
        
		htmlReport = new HTMLReport();
	}


	
	
	/**
	 * Reads the MasterSheet.xls to get all the executable files
	 * 
	 */

	@Test(priority = 1)
	public void getExecutableFiles() {

		resultSet = new LinkedHashMap<String, ArrayList<String>>();

		String requiredColHeader1 = "Test";
		String requiredColHeader2 = "Cadence";

		DriverScript main= new DriverScript();
		executableFiles = new ArrayList();

		// Create the MasterFile to read each of the Sub files */
		File masterFile = main.findFile(inputDirectory, masterFileName);
		if (masterFile == null) {
			appLogs.error("MasterSheet Not found in the specfied Directory : "
					+ inputDirectory);
			System.exit(1);
		}

		FileInputStream inputStream = null;

		try {
			inputStream = new FileInputStream(masterFile);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		Workbook xlReader = POIObjectCreator(masterFile, masterFile.toString(),
				inputStream);
		Sheet sheet = xlReader.getSheet(configSheet);
		executableFiles = getDataBelowCol(sheet, requiredColHeader1,
				requiredColHeader2, runMode);

		appLogs.info("Total Executable Files in MasterSheet are "
				+ executableFiles.size() + ". These are :");

		for (String fileName : executableFiles) {
			appLogs.info(fileName);
		}

	}
	
	
	/**
	 * Reads each executable file in the Mastersheet.xls to get all
	 * the executable subSheets and predefined variables in each file.	 
	 */

	@Test(priority = 1)
	public void getExecutableSheets() {

		String requiredCol1 = "Test";
		String requiredCol2 = "Variables";

		for (String executableFileName : executableFiles) {
			appLogs.info("");
			appLogs.info("");
			appLogs.info("<------------------------ Currenty Executing Sheet : "
					+ executableFileName
					+ " by Thread "
					+ Thread.currentThread() + " -------------------------->");
			appLogs.info("");

			LinkedHashMap<String, String> variables = new LinkedHashMap<String, String>();
			executableSheets = new ArrayList<String>();
			xlReaderObjs = new HashMap<Workbook, ArrayList<String>>();
			File executableFile = findFile(inputDirectory, executableFileName);
			if (executableFile == null) {
				appLogs.error("The file Name "+ executableFileName + " that is mentioned in the Master Sheet not found in the specfied Directory"
						+ " : " + inputDirectory);
				createResultSet(executableFileName, "N/A", "FAIL","File Not Found in the Directory", "N/A");
				// add here.
				continue;
			}

			int subSheetRow = 0;
			int variableRow = 0;

			FileInputStream inputStream = null;
			try {
				inputStream = new FileInputStream(executableFile);
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

			Workbook xlReader = POIObjectCreator(executableFile,executableFileName, inputStream);
			Sheet sheet = xlReader.getSheetAt(0);
			int totalRows = sheet.getPhysicalNumberOfRows();
			appLogs.debug("Total Number of Physical Rows in "
					+ executableFileName + " : " + totalRows);

			for (Row row : sheet) {
				for (Cell cell : row) {
	                 cell.setCellType(Cell.CELL_TYPE_STRING);
					if (cell.getStringCellValue().equalsIgnoreCase(requiredCol1)) {
						subSheetRow = cell.getRowIndex() + 1;
						appLogs.info("Sub Sheet row is at Row Number "+ subSheetRow);
					}
					if (cell.getStringCellValue().equalsIgnoreCase(requiredCol2)) {
						variableRow = cell.getRowIndex() + 1;
						appLogs.info("Variable Sheet row is at Row Number "+ variableRow);
					}
				}
			}

			for (int i = subSheetRow; i < variableRow - 1; i++) {
				 Row row2 = sheet.getRow(i);
				 if (row2 != null) {
					Cell cell1 = row2.getCell(0);
					Cell cell2 = row2.getCell(1);
					if (cell1 != null && cell2 != null) {
						if ((cell2.getStringCellValue().equalsIgnoreCase(runMode))&& (!cell1.getStringCellValue().equals(" "))) {
							executableSheets.add(cell1.getStringCellValue());
						}
					}

				}

			}

			for (int j = variableRow ; j < totalRows; j++) {
				Row row2 = sheet.getRow(j);
				if (row2 != null) {
					Cell cell1 = row2.getCell(0);
					Cell cell2 = row2.getCell(1);
					if (cell1 != null && cell2 != null) {
						if (!cell1.getStringCellValue().isEmpty()) {
							variables.put(cell1.getStringCellValue(),
									cell2.getStringCellValue());
						}

					}

				}

			}

			/* display subsheet names */
			appLogs.info("Total Executable Sheets in " + executableFileName+ " : " + executableSheets.size() + ".These are : ");
			if (executableSheets.size() > 0) {
				for (Object o : executableSheets) {
					appLogs.info(" -> " + o.toString());
				}
			}

			appLogs.info("Total variables in " + executableFileName + " : "+ variables.size() + ".These are : ");
			if (variables.size() > 0) {
				for (String variableKey : variables.keySet()) {
					appLogs.info(variableKey + " -> "+ variables.get(variableKey));
				}
			}

			/* create datastructures */
			fileVariables.put(executableFileName, variables);
			xlReaderObjs.put(xlReader, executableSheets);
			allXlReaderObjs.put(executableFileName, xlReaderObjs);

			// appLogs.info("Total Number of Variables in the config sheet in "
			// + fileName + " are - " + variables.size());
			// appLogs.info("-----------------------");
			appLogs.debug("Variables in" + executableFileName + " are -> "+ fileVariables.get(executableFileName));
			appLogs.debug("Total Sheets in " + executableFileName + " are -> "+ allXlReaderObjs.get(executableFileName).get(xlReader).size());

		}

	}// End of getExecutableSheets method
	
	
	
	/**
	 * Reads each executable Sub Sheet in Each executable File
	 * to store the each rows in a matrix.
	 * 
	 */
	@Test(priority = 2)
	public void readExecuatbleSheets() {

		int browserRow = 0;
		int nameIndex = 0;
		int actionIndex = 0;
		int locatorTypeIndex = 0;
		int locatorValueIndex = 0;
		int dataIndex = 0;

		int status = -1;

		first: for (String keys : allXlReaderObjs.keySet()) {
			String fileName = keys;

			second: for (Workbook keys2 : allXlReaderObjs.get(keys).keySet()) {
				appLogs.debug("Exact Sheets in file " + keys + " are : "+ allXlReaderObjs.get(keys).get(keys2));
				String browser = "";
				browser = getValueAfterCell(keys2.getSheetAt(0), "Browser");
				appLogs.debug("Run file " + fileName + " in " + browser + " browser.");
				if (browser.equals("")) {
					appLogs.error("Browser is not mentioned in the file : "+ fileName);
					// add here
					continue second;
				} else {
					appLogs.debug("Browser name is displayed");
				}

				third: 
					for (int i = 0; i < allXlReaderObjs.get(keys).get(keys2).size(); i++) {
					Workbook sheetReader = keys2;
					String sheetName = allXlReaderObjs.get(keys).get(keys2).get(i);
					Sheet currentSheet = sheetReader.getSheet(sheetName);

					if (!verifySheetPresence(sheetReader, sheetName)) {
						appLogs.error("Execuction failed : " + sheetName+ " is NOT present in the file : " + fileName);
						createResultSet(fileName, sheetName, "FAIL","Sheet is not Present in the File", "N/A");
						continue third;
					}

					appLogs.debug("Column Header Index position in : "+ currentSheet.getSheetName());

					/*
					 * find the index position of each mandatory header in the
					 * main folder
					 */
					nameIndex = columnHeaderIndex(currentSheet, "Name");
					actionIndex = columnHeaderIndex(currentSheet, "Action");
					locatorTypeIndex = columnHeaderIndex(currentSheet,"Locator Type");
					locatorValueIndex = columnHeaderIndex(currentSheet,"Locator Value");
					dataIndex = columnHeaderIndex(currentSheet,"Test Data/Options");

					appLogs.debug("Name header is at index position - "+ nameIndex);
					appLogs.debug("Action header is at index position - "+ actionIndex);
					appLogs.debug("Locator Type header is at index position - "+ locatorTypeIndex);
					appLogs.debug("Locator Header header is at index position - "+ locatorValueIndex);
					appLogs.debug("Test Data/Options header is at index position - "+ dataIndex);

					if (nameIndex == 0 || actionIndex == 0 || locatorTypeIndex == 0 || dataIndex == 0 || locatorValueIndex == 0) {
						appLogs.error("Exection Failed : One or more mandatory column header is missing in sheet : "
					                  + sheetName + " and file : " + fileName);
						createResultSet(fileName, sheetName, "FAIL","Required column headers are missing.", "N/A");
						continue third;

					} else {
						appLogs.debug("All the mandatory column hearder are present in sheet : "
							    + sheetName + " and file : " + fileName);
					}

					/*
					 * Verify that the last Action should be Quit. Else stop
					 * execution of the current sheet.
					 */
					int actualPhysicalRows = getPhysicalRows(sheetReader,sheetName, actionIndex, "mainSheet");
					appLogs.info("Total Number of Steps in " + sheetName+ "are : " + actualPhysicalRows);
					if (actualPhysicalRows == 0) {
						appLogs.error("There is no Quit Method in the current Sheet : "
								  + sheetName + " in the file " + fileName);
						createResultSet(fileName, sheetName, "FAIL",
								  "There is no quit method at the end.", "N/A");
						continue third; // Stop execution of the current Sheet as it has no Quit at the end.
					}

					else {
						appLogs.debug("There is Quit Method in the current Sheet : "
								+ sheetName + " in the file " + fileName);
					}

					/*
					 * Create a hashmap of all the important Column data, one
					 * row at a time
					 */
					appLogs.debug("Sheet " + sheetName+ " has following data in the Action Column");
					int currentlyExecutingRowInMain = 1;
					sheetRowData = new ArrayList<ArrayList<String>>();
					for (int j = 1; j <= actualPhysicalRows; j++) { // starting hashmap creation for loop
																	 
						/* Clear row data for each row */

						rowData = new ArrayList<String>();

						rowData.add("Main Sheet");
						rowData.add(getCellData(sheetReader, sheetName, j,nameIndex));
						rowData.add(getCellData(sheetReader, sheetName, j,actionIndex));
						rowData.add(getCellData(sheetReader, sheetName, j,locatorTypeIndex));
						rowData.add(getCellData(sheetReader, sheetName, j,locatorValueIndex));
						rowData.add(getCellData(sheetReader, sheetName, j,dataIndex));
						rowData.add(Integer.toString(currentlyExecutingRowInMain));
						rowData.add("N/A");

						appLogs.debug("SheetName = " + rowData.get(0)
								+ " ---- " + " Name = " + rowData.get(1)
								+ " ---- " + "Action = " + rowData.get(2)
								+ " ---- " + "Locator Type = " + rowData.get(3)
								+ " ---- " + "Locator Value = "
								+ rowData.get(4) + " ---- "
								+ "Test Data/Options = " + rowData.get(5));

						/* Verify Location of the Action item */
						File commonFile = findFile(commonInputDirectory,
								rowData.get(2).toString());

						if (commonFile != null) {
							
							if(!rowData.get(5).equals("N/A")){
								int state= createCommonSheetVariables(fileName, sheetName, commonFile.getName(), currentlyExecutingRowInMain, rowData.get(5).split(","));
								if(state==0){
									appLogs.error("Configuration Issue at Row - " + currentlyExecutingRowInMain
											+ ":: Incorrect Format for the variabes provided for Common Sheet "
											+ commonFile.getName() + " in Sheet - " + sheetName);
									createResultSet(fileName, sheetName, "FAIL","Configuration Issue at Row - " + currentlyExecutingRowInMain
											+ " :: Incorrect Format for the variables provided for Common Sheet "
											+ commonFile.getName() + " in Sheet - " + sheetName, "N/A");
									continue third; // stop execution of the current sheet due to incorrect variable intialization for common sheet	
								}
							}
					
							String commonFileName = rowData.get(2);
							rowData.clear();
							int commonNameIndex = 0;
							int commonActionIndex = 0;
							int commonLocatorTypeIndex = 0;
							int commonLocatorValueIndex = 0;
							int commonDataIndex = 0;

							appLogs.info("Execute sheet " + commonFileName+ " from the common folder");
							
							/* add logic to generate row Data HashMap from the Common Sheet*/
							
							FileInputStream commonInputStream = null;
							try {
								commonInputStream = new FileInputStream(
										commonFile);
							} catch (FileNotFoundException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
							Workbook commonXlReader = POIObjectCreator(commonFile, commonFileName,commonInputStream);
							Sheet commonSheet = commonXlReader.getSheetAt(0);

							
							 /* find the index position of each mandatory column header in the common file*/
							 
							commonNameIndex = columnHeaderIndex(commonSheet,"Name");
							commonActionIndex = columnHeaderIndex(commonSheet,"Action");
							commonLocatorTypeIndex = columnHeaderIndex(commonSheet, "Locator Type");
							commonLocatorValueIndex = columnHeaderIndex(commonSheet, "Locator Value");
							commonDataIndex = columnHeaderIndex(commonSheet,"Test Data/Options");

							/*
							 * appLogs.info(
							 * "Common Name header is at index position - " +
							 * commonNameIndex); appLogs.info(
							 * "Common Action header is at index position - " +
							 * commonActionIndex); appLogs.info(
							 * "Common Locator Type header is at index position - "
							 * + commonLocatorTypeIndex); appLogs.info(
							 * "Common Locator Header header is at index position - "
							 * + commonLocatorValueIndex);
							 */

							/* int actualPhysicalRowsInCommon = getPhysicalRows(commonXlReader,commonXlReader.getSheetName(0).toString() ,
							          commonActionIndex , "commonSheet");
							 appLogs.info("first sheet name in common file is "
							 +commonXlReader.getSheetName(0)+ " it has " +
							  actualPhysicalRowsInCommon + " rows." ); */
							appLogs.info("Total rows in common sheet" + commonSheet.getPhysicalNumberOfRows());
							
							int currentlyExecutingRowInCommon = 1;
							for (int k = 1; k < commonSheet
									.getPhysicalNumberOfRows(); k++) {

								commonSheetRowData = new ArrayList<String>();
								// adding2D = new
								// ArrayList<ArrayList<String>>();
								commonSheetRowData.add(commonFile.getName());
								commonSheetRowData.add(getCellData(commonXlReader,commonXlReader.getSheetName(0), k,commonNameIndex));
								commonSheetRowData.add(getCellData(commonXlReader,commonXlReader.getSheetName(0), k,commonActionIndex));
								commonSheetRowData.add(getCellData(commonXlReader,commonXlReader.getSheetName(0), k,commonLocatorTypeIndex));
								commonSheetRowData.add(getCellData(commonXlReader,commonXlReader.getSheetName(0), k,commonLocatorValueIndex));
								commonSheetRowData.add(getCellData(commonXlReader,commonXlReader.getSheetName(0), k,commonDataIndex));
								commonSheetRowData.add(Integer.toString(currentlyExecutingRowInCommon));
								commonSheetRowData.add(Integer.toString(currentlyExecutingRowInMain));

								appLogs.debug("SheetName = "
										+ commonSheetRowData.get(0) + " ---- "
										+ " Name = "
										+ commonSheetRowData.get(1) + " ---- "
										+ "Action = "
										+ commonSheetRowData.get(2) + " ---- "
										+ "Locator Type = "
										+ commonSheetRowData.get(3) + " ---- "
										+ "Locator Value = "
										+ commonSheetRowData.get(4) + " ---- "
										+ "Test Data/Options = "
										+ commonSheetRowData.get(4));
								sheetRowData.add(commonSheetRowData);

								
								currentlyExecutingRowInCommon++;
							}
							

							appLogs.debug("---------------------- Execution moves back to the Main sheet from the common folder -------------------");

						} else {
							appLogs.debug("Execute action " + rowData.get(2)+ " from the Action class Methods");
							sheetRowData.add(rowData);

						}

						sheetData.put(fileName + ":" + sheetName + ":"+ browser, sheetRowData);

						currentlyExecutingRowInMain++;

					}// End of hashmap creating for loop.
				}

			}
		}

	}// readExecutableSheet method ends here.

	
	
	
	/**
	 * Displays all the input data, that is, Total Executable Files,
	 * Total Executable Sheets in each file, Cell data in each executable file.
	 */
	@Test(priority = 3)
	public void displayInput() {

		appLogs.info("Total number of Executable Files :"
				+ executableFiles.size());
		appLogs.info("Total executable sheets in each file are");

		for (String keys : allXlReaderObjs.keySet()) {
			for (Workbook keys2 : allXlReaderObjs.get(keys).keySet()) {

				appLogs.info("Executable Sheets in file " + keys + " are : "
						+ allXlReaderObjs.get(keys).get(keys2));
			}
		}

		appLogs.info("Total Size of the Sheet Data Metrix " + sheetData.size()
				    + ". Rows in each sheet are: ");
		for (String keys : sheetData.keySet()) {

			appLogs.info(keys + "--> " + sheetData.get(keys));
		}

		for (String executableFileName : fileVariables.keySet()) {
			appLogs.info("Total Variables in the File " + executableFileName
					+ " are: " + fileVariables.get(executableFileName));
		}
		
		
		appLogs.info ("Total common sheet variables are : ");
		
		for(String keys : allCommonSheetVariables.keySet()){
			
			appLogs.info(keys + " : " + allCommonSheetVariables.get(keys));
		}	
		

	}

	
	/**
	 * Invokes the executer method defined in Utils class
	 * for each executable sheet. 
	 */

	@Test(priority = 4)
	public void invokeExecutor() {

		/*ExecutorService service = Executors.newFixedThreadPool(1);
		 try{
			 for (String fileKey : sheetData.keySet()){
				 final String key= fileKey;
				 final String[] sheetInfo = fileKey.split(":" , 3);
				 service.submit(new Runnable() {
			            public void run() {
			            
			            	new Utils().executor(sheetInfo, sheetData.get(key),
									fileVariables.get(sheetInfo[0]));
			            }
			        });
			 }
		 }
		 finally{
			 service.shutdown();
		 } */
		
		//appLogs.info("Currently Running thread : "  +Thread.currentThread().getName());
		// create executor service over here.
		for (String fileKey : sheetData.keySet()) {    
			    	
     		    	Utils utils = new Utils();
					String[] sheetInfo =fileKey.split(":", 3);
					appLogs.debug("first String - " + sheetInfo[0]);
					appLogs.debug("Second String - " + sheetInfo[1]);
					appLogs.debug("Third String - " + sheetInfo[2]);
					utils.executor(sheetInfo, sheetData.get(fileKey),					
							 fileVariables.get(sheetInfo[0]), allCommonSheetVariables);
		} 
		
	}
	

	/**
	 * Creates the HTML Report 
	 * 
	 */
	@Test(priority = 6)
	public void createHTMLResult() {

		//test
		htmlReport.checkResultDir();
		htmlReport.makeHtmlReport(resultSet, executableFiles);
	}
	

	/**
	 * Displays the final output in the console.
	 * 
	 */
	@Test(priority = 5)
	public void displayOutput() {

		appLogs.info("Final Execution Result");
		for (String fileName : resultSet.keySet()) {
			appLogs.info(fileName + " : " + resultSet.get(fileName));
		}

	}

	
	/**
	 * Displays the final output in the console.
	 * 
	 */
	public int createCommonSheetVariables(String fileName, String sheetName, String commonSheetName, int rowNumber, String [] splitVariables){
		commonSheetVariables = new HashMap<String, String>();
		
		appLogs.info("Total variables: " + splitVariables.length);
		
		 for(int i=0; i<splitVariables.length; i++){
			   String [] splitKeyValuePair = splitVariables[i].split("::");
			   appLogs.info("Total 2: " + splitKeyValuePair.length);
			   appLogs.info("Key " + splitKeyValuePair[0]);
			   appLogs.info("Value " + splitKeyValuePair[1]);
			     if(splitKeyValuePair.length==2 && splitKeyValuePair[0].toString().trim().substring(0, 5).equals("var <")){
			    	 String key = splitKeyValuePair[0].toString().trim().substring(3);
			    	 String value = splitKeyValuePair[1].trim();		    	 
     		    	  commonSheetVariables.put(key.trim(), value.trim());
			     }
			     else{
			    	 return 0;
			     }
		 }
		 
		 allCommonSheetVariables.put(fileName +":" +sheetName +":" +commonSheetName+ ":" +rowNumber, commonSheetVariables);
		 return 1;
	}

	/**
	 * Adds the execution results for each sheet in resultSet hashMap
	 * 
	 * @param fileName -  Name of the currently executing file
	 * @param SheetName - Name of the current executing Sheet
	 * @param status - 'PASS' or 'FAIL'
	 * @param exceptionMessage - The exception message that would have been raised.
	 * @param screenshotPath - The screenshotName in case of any exception
	 */
	public void createResultSet(String FileName, String sheetName,
			String status, String exceptionMessage, String screenshotPath) {
		appLogs.debug("Inside createResultSet Method");
		subResultSet = new ArrayList<String>();
		// subResultSet.add(sheetName);
		subResultSet.add(status);
		subResultSet.add(exceptionMessage);
		subResultSet.add(screenshotPath);

		resultSet.put(FileName + ":" + sheetName, subResultSet);
	}

	
	/**
	 * Checks if the file is present in the current directory.
	 * 
	 * @param directory -  the path where the file has be searched in
	 * @param fileName - The file name to be searched
	 * @return File object if the file is found, else NULL
	 */
	public File findFile(String directory, String fileName) {
		String name = fileName;
		File f = new File(directory);
		File[] listOfFiles = f.listFiles();

		for (File file : listOfFiles) {
			/*
			 * check if the file with specific name is present in the directory.
			 * Also ignores the version of file that is already opened
			 */
			if (file.getName().contains(fileName)
					&& (!file.getName().contains("~lock"))) {
				appLogs.debug("File " + fileName + " found at " + file);
				return file;
			}

		}
		return null;

	}

}
