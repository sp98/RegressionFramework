package core;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.concurrent.BrokenBarrierException;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.CyclicBarrier;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.testng.annotations.Test;

import reporting.HTMLReport;
import utilities.Utils;
import fileReader.XLReading;



public class DriverScript extends XLReading  {

	
	private String inputDirectory = null;   //Input directory where are all the files need to be placed.
	private String commonInputDirectory = null; //Common Directory inside Input directory where commmon sheets are placed.
	private String masterFileName = "MasterSheet"; //Name of the MasterSheet to read all the executable sheets.
	private String configSheet = "config";
	private String runMode = "";  // Daily/Weekly/Release Run mode. 
	private int thread_Count= 0;
	public static boolean enable_grid= false;
	private static String raw_grid_config = null;
	private boolean gridConfigCreationStatus = false;

	/*Result Set Data Structure*/
	private static LinkedHashMap<String, ArrayList<String>> resultSet = null;
	private ArrayList<String> subResultSet = null;
	
	
	private ArrayList<String> executableFiles = null; // List of all the Executable Files present in the MasterSheet.
	
	private HashMap<String, LinkedHashMap<String, String>> fileVariables = null;  //Collection of all the Variable Names in one file. 
	private ArrayList<String> executableSheets = null;
	private HashMap<String, HashMap<Workbook, ArrayList<String>>> allXlReaderObjs = null;
	private HashMap<Workbook, ArrayList<String>> xlReaderObjs = null;

	/* Matrix for all the row data in a each sheet */
	private LinkedHashMap<String, ArrayList<ArrayList<String>>> sheetData = null;
	private ArrayList<ArrayList<String>> sheetRowData = null;
	private ArrayList<String> rowData = null;
	private ArrayList<String> commonSheetRowData = null;
	
	
	/*Hashmap that contains the runtime variables created in the common Sheet */
	private LinkedHashMap<String , HashMap<String, String>> allCommonSheetVariables = null;
	HashMap<String, String> commonSheetVariables= null;

	/*ArrayList of ArrayList to Store the Grid config*/
	private ArrayList<ArrayList<String>> finalGridConfig = null;
	private ArrayList<String> subGridConfig = null;
	
	private HTMLReport htmlReport = null;   // A reference to the HTMLReport class in the 'reporting' package.


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
		finalGridConfig = new ArrayList<ArrayList<String>>();
		
		htmlReport = new HTMLReport();  //initialize the object for HTMLReport class present in 'reporting' package
	}


	
	
	/**
	 * Reads the MasterSheet.xls to get all the executable files
	 * 
	 */

	@Test(priority = 1)
	public void getExecutableFiles() {

		resultSet = new LinkedHashMap<String, ArrayList<String>>();     //Initialize the ResultSet Linked HashMap to store the results.
        
		/* Mandatory column headers in MasterSheet */
		String requiredColHeader1 = "Test";
		String requiredColHeader2 = "Cadence";

		DriverScript main= new DriverScript();
		executableFiles = new ArrayList();

		// Check the MasterFile to read each of the Sub files */
		File masterFile = main.findFile(inputDirectory, masterFileName);
		if (masterFile == null) {
			appLogs.error("MasterSheet Not found in the specfied Directory : "
					+ inputDirectory);   //MasterFile is not present in the specified directory.
			System.exit(1);
		}

		FileInputStream inputStream = null;

		try {
			inputStream = new FileInputStream(masterFile);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
       
		/*create the Workbook object to read data from the masterfile */
		Workbook xlReader = POIObjectCreator(masterFile, masterFile.toString(),inputStream);
		Sheet sheet = xlReader.getSheet(configSheet);
		executableFiles = getDataBelowCol(sheet, requiredColHeader1,requiredColHeader2, runMode);
		
		/*Set the Thread count from the MasterSheet config */
		try{
			if(Integer.parseInt(getValueAfterCell(xlReader.getSheetAt(0), "Num_Threads"))>0)
				thread_Count= Integer.parseInt(getValueAfterCell(xlReader.getSheetAt(0), "Num_Threads"));	//set the thread count value.
			else
				thread_Count = 2;   //keep the default thread count value to 2 if nothing is specified.
		}
		catch(Exception e){
			appLogs.info("Execption Raised : Incorrect Thread count Provided in the Master sheet- "
		       + e.getMessage());		
		}
			
		appLogs.debug("Total Thread Count : " + thread_Count);
		
		/*Create Grid config if Enabled*/
		if(getValueAfterCell(xlReader.getSheetAt(0), "Enable-Grid").equalsIgnoreCase("YES")){
			enable_grid=true;
			raw_grid_config = getValueAfterCell(xlReader.getSheetAt(0), "OS|NodeURL|Browsers");
			gridConfigCreationStatus = createGridConfig(raw_grid_config);
		}
		
		  appLogs.info("Is Grid Enabled?" + "--> " +enable_grid);
		  appLogs.info("Grid config" + "--> " +raw_grid_config);
		  appLogs.info("Grid Config Creation Status" + "-->" +gridConfigCreationStatus );
		  appLogs.info("Final Grid Config is" + "-->" +finalGridConfig);
		}
	

	/**
	 * Reads each executable file in the Mastersheet.xls to get all
	 * the executable subSheets and predefined variables in each file.	 
	 */

	@Test(priority = 1)
	public void getExecutableSheets() {
 
		/*Mandatory column headers in config sheet of each file */
		String requiredCol1 = "Test";
		String requiredCol2 = "Variables";

		for (String executableFileName : executableFiles) {
			
			LinkedHashMap<String, String> variables = new LinkedHashMap<String, String>();
			executableSheets = new ArrayList<String>();
			xlReaderObjs = new HashMap<Workbook, ArrayList<String>>();
			File executableFile = findFile(inputDirectory, executableFileName);
			
			/* Check if the file name mentioned in the MasterSheet is present in the specified Directory */
			if (executableFile == null) {
				appLogs.error("The file Name "+ executableFileName + " that is mentioned in the Master Sheet not found in the specfied Directory"
						+ " : " + inputDirectory);    //File not present in the specified directory. 
				createResultSet(executableFileName, "N/A", "FAIL","File Not Found in the Directory", "N/A");
				continue;
			}

			/*Initializing the index position of the Sheet and 'Variable' column header to 0 */ 
			int subSheetRow = 0;
			int variableRow = 0;

			
			FileInputStream inputStream = null;     //create a new InputStream Object to interact with the file.
			try {
				inputStream = new FileInputStream(executableFile);
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
 
			/*Create the Workbook object for the curretly executing excel file */
			Workbook xlReader = POIObjectCreator(executableFile,executableFileName, inputStream);
			Sheet sheet = xlReader.getSheetAt(0);      //Read data from the first Sheet
			int totalRows = sheet.getPhysicalNumberOfRows();
			appLogs.debug("Total Number of Physical Rows in "
					+ executableFileName + " : " + totalRows);

			/* Find the exact index of the Sheet and Variable in the config sheet */
			for (Row row : sheet) {
				for (Cell cell : row) {
	                 cell.setCellType(Cell.CELL_TYPE_STRING);
					if (cell.getStringCellValue().equalsIgnoreCase(requiredCol1)) {
						subSheetRow = cell.getRowIndex() + 1;
						appLogs.debug("Sub Sheet row is at Row Number "+ subSheetRow);
					}
					if (cell.getStringCellValue().equalsIgnoreCase(requiredCol2)) {
						variableRow = cell.getRowIndex() + 1;
						appLogs.debug("Variable Sheet row is at Row Number "+ variableRow);
					}
				}
			}

			/* Loop to find all the executable sheets in the config and add them to'executableSheets' data structure */
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

			/* Loop to find all the predefined variables and add them to 'variables' hashMap */
			for (int j = variableRow ; j < totalRows; j++) {
				Row row2 = sheet.getRow(j);
				if (row2 != null) {
					Cell cell1 = row2.getCell(0);
					Cell cell2 = row2.getCell(1);
					if (cell1 != null && cell2 != null) {
						if (!cell1.getStringCellValue().isEmpty()) {
							variables.put(cell1.getStringCellValue(),   
									cell2.getStringCellValue());                   // add each variable in the 'variables' hashMap
						}

					}

				}

			}

			/* display subsheet names in the console*/
			/*appLogs.info("Total Executable Sheets in " + executableFileName+ " : " + executableSheets.size() + ".These are : ");
			if (executableSheets.size() > 0) {
				for (Object o : executableSheets) {
					appLogs.info(" -> " + o.toString());
				}
			}*/

			/* Display all the predefined variables in sheet */
			/*appLogs.info("Total variables in " + executableFileName + " : "+ variables.size() + ".These are : ");
			if (variables.size() > 0) {
				for (String variableKey : variables.keySet()) {
					appLogs.info(variableKey + " -> "+ variables.get(variableKey));
				}
			}*/

			/* create data structures */
			fileVariables.put(executableFileName, variables);  // Add all variables hashmap to the fileVariables hashmap for each executable File
			xlReaderObjs.put(xlReader, executableSheets); //Add all the executable sheets in a file to xlReaderObjs hashmap for each workbook object.
			allXlReaderObjs.put(executableFileName, xlReaderObjs); //All all the xlReadersObjs hashmap to 'allXlReaderObjs' hashmap.

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

		/*initializing the local column header variables */
		int browserRow = 0;
		int nameIndex = 0;
		int actionIndex = 0;
		int locatorTypeIndex = 0;
		int locatorValueIndex = 0;
		int dataIndex = 0;

		int status = -1;

		 first:   //First For Loop label
			 /*Loop through all the Workbook objects for each execuatable sheet */
			for (String keys : allXlReaderObjs.keySet()) {
			String fileName = keys;

			second: //Second for Loop Label
				/*Loop through all the executable Sheets using there WorkBook Objects */
				for (Workbook keys2 : allXlReaderObjs.get(keys).keySet()) {
				  appLogs.debug("Exact Sheets in file " + keys + " are : "+ allXlReaderObjs.get(keys).get(keys2));
				  
				  String browser = "";  //Declare the String variable browser.
				  if(!enable_grid){     //Initialize the Browser from individual sheet only of the Grid is not enabled.
					  browser = getValueAfterCell(keys2.getSheetAt(0), "Browser");   //Get the value of the browser value from the config sheet.
					  appLogs.debug("Run file " + fileName + " in " + browser + " browser.");  
				  }
				                   
				 

				third: // Third for Loop Label.
					
					for (int i = 0; i < allXlReaderObjs.get(keys).get(keys2).size(); i++) {
					  Workbook sheetReader = keys2;
					  String sheetName = allXlReaderObjs.get(keys).get(keys2).get(i);
					  Sheet currentSheet = sheetReader.getSheet(sheetName);

					/* check if the sheetname provided in the config is actually present in the file */
					if (!verifySheetPresence(sheetReader, sheetName)) {
						appLogs.error("Execuction failed : " + sheetName+ " is NOT present in the file : " + fileName);
						createResultSet(fileName, sheetName, "FAIL","Sheet is not Present in the File", "N/A");
						continue third;
					}

					appLogs.debug("Column Header Index position in : "+ currentSheet.getSheetName());

					/*find the index position of each mandatory header in the main folder*/
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

					/* Check if the any of the Mandatory column Headers are missing */
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
					appLogs.debug("Total Number of Steps in " + sheetName+ "are : " + actualPhysicalRows);
					if (actualPhysicalRows == 0) {
						appLogs.error("There is no Quit Method in the current Sheet : "
								  + sheetName + " in the file " + fileName);
						createResultSet(fileName, sheetName, "FAIL","There is no quit method at the end.", "N/A");
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

						/* Add data for the current Row in the rowData */
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

						/* Check if the current action is in common sheet */
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

							appLogs.debug("Execute sheet " + commonFileName+ " from the common folder");
							
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
							appLogs.debug("Total rows in common sheet" + commonSheet.getPhysicalNumberOfRows());
							
							int currentlyExecutingRowInCommon = 1;
							for (int k = 1; k < commonSheet
									.getPhysicalNumberOfRows(); k++) {

								commonSheetRowData = new ArrayList<String>();
								// adding2D = new
								// ArrayList<ArrayList<String>>();
								
								/* Add data for each row in the common sheet in the commonSheetRow Data */
								commonSheetRowData.add(commonFile.getName());
								commonSheetRowData.add(getCellData(commonXlReader,commonXlReader.getSheetName(0), k,commonNameIndex));
								commonSheetRowData.add(getCellData(commonXlReader,commonXlReader.getSheetName(0), k,commonActionIndex));
								commonSheetRowData.add(getCellData(commonXlReader,commonXlReader.getSheetName(0), k,commonLocatorTypeIndex));
								commonSheetRowData.add(getCellData(commonXlReader,commonXlReader.getSheetName(0), k,commonLocatorValueIndex));
								commonSheetRowData.add(getCellData(commonXlReader,commonXlReader.getSheetName(0), k,commonDataIndex));
								commonSheetRowData.add(Integer.toString(currentlyExecutingRowInCommon));
								commonSheetRowData.add(Integer.toString(currentlyExecutingRowInMain));

								appLogs.debug("SheetName = " + commonSheetRowData.get(0) + " ---- "
										+ " Name = " + commonSheetRowData.get(1) + " ---- "
										+ "Action = " + commonSheetRowData.get(2) + " ---- "
										+ "Locator Type = " + commonSheetRowData.get(3) + " ---- "
										+ "Locator Value = " + commonSheetRowData.get(4) + " ---- "
										+ "Test Data/Options = " + commonSheetRowData.get(4));
								
								sheetRowData.add(commonSheetRowData);  //Add data for each row into 'SheetRowData' arrayList

								
								currentlyExecutingRowInCommon++;    
							}
							

							appLogs.debug("---------------------- Execution moves back to the Main sheet from the common folder -------------------");

						} else {
							appLogs.debug("Execute action " + rowData.get(2)+ " from the Action class Methods");
							sheetRowData.add(rowData);

						}

						/*This hashMap helps to keep a track of all the Data in one sheet
						 * Here, the key for each entry - 'File Name : Sheet Name : Browser' has all 
						 * the data for the corresponding sheet. 
						 */
						if(!enable_grid){
							sheetData.put(fileName + "::" + sheetName + "::"+ browser, sheetRowData);  //Add Complete Sheet Data to 'SheetData' hashmap.	
						}
						else{
							for(ArrayList<String> subConfig :finalGridConfig){
								
								for(int m=2; m<=subConfig.size()-1;m++){
									sheetData.put(fileName + "::" + sheetName +"::" + subConfig.get(0)+ "::" 
											+ subConfig.get(1) + "::" + subConfig.get(m) , sheetRowData);
								}
							}
							
						}
						

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
		
		appLogs.info(" ");
		appLogs.info(" <---------------------------- INPUT DATA ----------------------------------> ");
		appLogs.info(" ");
		
		/* Print all the Executable files in the console */
		appLogs.info("Total Executable Files in MasterSheet are "
				   + executableFiles.size() + ". These are :");
		
		for (String fileName : executableFiles) {
			appLogs.info(fileName);
		} 
		
		appLogs.info(" ");
      
		//appLogs.info("Total number of Executable Files :"+ executableFiles.size());
		//appLogs.info("Total executable sheets in each file are :");

		/* Print the total Executable sheets in each file */
		for (String keys : allXlReaderObjs.keySet()) {
			for (Workbook keys2 : allXlReaderObjs.get(keys).keySet()) {

				appLogs.info("Executable Sheets in file " + keys + " are " 
						 + allXlReaderObjs.get(keys).get(keys2).size() + " : "
						+ allXlReaderObjs.get(keys).get(keys2));
			}
		}
		
		appLogs.info(" ");
		
       /*Print the data present in each sheet */
		appLogs.info("Data in each Sheet : ");
		for (String keys : sheetData.keySet()) {
			appLogs.info(keys + "--> " + sheetData.get(keys));
		}
		
		appLogs.info(" ");

		/*Print the global variables declared in each file*/
		for (String executableFileName : fileVariables.keySet()) {
			appLogs.info("Total Variables in the File " + executableFileName
					+ " are: " + fileVariables.get(executableFileName));
		}
		
		appLogs.info(" ");
		
		/*Print the total common sheet variables in each sheet */
		appLogs.info ("Total common sheet variables are : ");		
		for(String keys : allCommonSheetVariables.keySet()){		
			appLogs.info(keys + " : " + allCommonSheetVariables.get(keys));
		}	
		
		appLogs.info(" ");
		
		appLogs.info("Total Thread Count : " + thread_Count);
		appLogs.info(" ");
	}

	
	/**
	 * Invokes the executer method defined in Utils class
	 * for each executable sheet. 
	 */

	@Test(priority = 4)
	public void invokeExecutor() {
	
		ArrayList<Thread> threadTracker = new ArrayList<Thread>();   //Create ArrayList to Keep Track of all the Treads that get created.

		
		final CountDownLatch latch = new CountDownLatch(sheetData.size());
		ExecutorService taskExecutor = Executors.newFixedThreadPool(thread_Count);
		
		for (final String fileKey : sheetData.keySet()) {
			
			final Utils utils = new Utils();
			final String[] sheetInfo =fileKey.split("::");
			taskExecutor.execute( new Runnable(){

				@Override
				public void run() {
					// TODO Auto-generated method stub
					utils.executor(sheetInfo, sheetData.get(fileKey),					
							 fileVariables.get(sheetInfo[0]), allCommonSheetVariables);
					
					 latch.countDown();
				}
				
			});
		}
		
		try {
	        latch.await();  // wait untill latch counted down to 0
	    } catch (InterruptedException e) {
	        // TODO Auto-generated catch block
	        e.printStackTrace();
	    }

		taskExecutor.shutdown();
		
		
	} 
	

	/**
	 * Creates the HTML Report 
	 * 
	 */
	@Test(priority = 6)
	public void createHTMLResult() {

		//test
		htmlReport.checkResultDir();
		if(!enable_grid)
		htmlReport.makeHtmlReport(resultSet, executableFiles);
		else
		htmlReport.makeGridHtmlReport(resultSet, executableFiles);
	}
	

	/**
	 * Displays the final output in the console.
	 * 
	 */
	@Test(priority = 5)
	public void displayOutput() {
		appLogs.info(" ");
		appLogs.info(" <------------- OUTPUT DATA -------------> ");
		appLogs.info(" ");
		appLogs.debug("Final Execution Result");
		for (String fileName : resultSet.keySet()) {
			appLogs.info(fileName + " : " + resultSet.get(fileName));
		}
		appLogs.info(" ");

	}

	
	/**
	 * Checks the 'Test Data/Options' for any dynamic variables and adds them to 
	 * commonSheetVariables HashMap
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param commonSheetName - Name of the common sheet that is being called in the current row.
	 * @param rowNumber - Current Executing row in the Main Sheet
	 * @param slitVariables - String array of Variables defined in the Test Data/Options sheet separated by ,
	 * @return 0 if the format of the variables defined is not correct 
	 */
	public int createCommonSheetVariables(String fileName, String sheetName, String commonSheetName, int rowNumber, String [] splitVariables){
		commonSheetVariables = new HashMap<String, String>();
		
		appLogs.debug("Total Common Sheet Variables in file : " + fileName + " are " + splitVariables.length +
				 ". These are : ");
		
		 for(int i=0; i<splitVariables.length; i++){
			   String [] splitKeyValuePair = splitVariables[i].split("::");
			   appLogs.debug("Total 2: " + splitKeyValuePair.length);
			   appLogs.debug("Key " + splitKeyValuePair[0]);
			   appLogs.debug("Value " + splitKeyValuePair[1]);
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
	public void createResultSet(String FileName, String sheetName, String status, String exceptionMessage, String screenshotPath) {
		appLogs.debug("Inside createResultSet Method");
		subResultSet = new ArrayList<String>();
		// subResultSet.add(sheetName);
		subResultSet.add(status);
		subResultSet.add(exceptionMessage);
		subResultSet.add(screenshotPath);

		resultSet.put(FileName + "::" + sheetName +"::" +"N/A"
				+ "::" +"N/A"+ "::" +"N/A", subResultSet);
	}
	
	/**
	 * Overloaded method that Adds the execution results for each sheet in resultSet hashMap
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param status - 'PASS' or 'FAIL'
	 * @param exceptionMessage - The exception message that would have been raised.
	 * @param screenshotPath - The screenshotName in case of any exception
	 */
	public void createResultSet(String[] sheetInfo, String status, String exceptionMessage, String screenshotPath) {
		appLogs.debug("Inside createResultSet Method");
		subResultSet = new ArrayList<String>();
		// subResultSet.add(sheetName);
		subResultSet.add(status);
		subResultSet.add(exceptionMessage);
		subResultSet.add(screenshotPath);

		//resultSet.put(FileName + ":" + sheetName, subResultSet);
		if(!enable_grid){
			resultSet.put(sheetInfo[0]+ "::" +sheetInfo[1]+ "::" +"N/A"
					+ "::" +"N/A"+ "::" +"N/A", subResultSet);
		}
		else{
			resultSet.put(sheetInfo[0]+ "::" +sheetInfo[1]+ "::" +sheetInfo[2]
					+ "::" +sheetInfo[3]+ "::" +sheetInfo[4], subResultSet);
		}
		
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
	
	public boolean createGridConfig(String raw_grid_config){
		
		if(raw_grid_config.equals("")){
			appLogs.error("Grid is enabled but no config info is provided");
			return false;
		}
		
	  String [] commaSplitedGridConfig = raw_grid_config.split(",");	
	  
	  for(String grid_config : commaSplitedGridConfig ){
		  String [] pipeSplitedGridConfig = grid_config.split("\\|");
		   if(pipeSplitedGridConfig.length<=2){
			   appLogs.info(pipeSplitedGridConfig.length);
			   return false;
		   }
		   else{
			   subGridConfig = new ArrayList<String>();
			   for(String config : pipeSplitedGridConfig){
				   subGridConfig.add(config.trim());
			   }
			   finalGridConfig.add(subGridConfig);
		   }
	  }
	  
	  return true;
	}

}
