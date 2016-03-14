package utilities;

import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.log4j.Logger;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Platform;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.htmlunit.HtmlUnitDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.phantomjs.PhantomJSDriver;
import org.openqa.selenium.phantomjs.PhantomJSDriverService;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.mysql.jdbc.Connection;
import com.mysql.jdbc.Statement;

import core.DriverScript;


public class Utils extends DriverScript {

	WebDriver driver = null;       //declaring WebDriver object.
	private Logger appLogs = null; //declaring Logger class Object for logging purpose
	WebDriverWait wait = null;     //
	Actions actions = null;        //declaring an Actions API class object
	DesiredCapabilities  caps = null;
	
	LinkedHashMap<String, HashMap<String, String>> allCommonSheetVars= null;    // declaring a linked HashMap variable
	//public static int parallelCounter = 0;  //counter used to control the parallel processing
	

	/**
	 * Class Constructor
	 * 
	 */
	public Utils() {
		// TODO Auto-generated constructor stub
		org.apache.log4j.PropertyConfigurator.configure(System
				.getProperty("user.dir") + "/log4j.properties");
		appLogs = Logger.getRootLogger();
	}
	
	/**
	 * Reads each row in the sheet and calls the appropriate method based on the 
	 * 'Action' column
	 * 
	 * @param sheetInfo -  String array which has
	 * @param sheetMatrix - complete data for a particular sheet.
	 * @param fileVariables - All the variables (specified in config and during run time) for the file
	 * @param allCommonSheetVariables - Variables created in front of the common sheet. 
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int  executor(String[] sheetInfo, ArrayList<ArrayList<String>> sheetMatrix,
			HashMap<String, String> fileVariables , LinkedHashMap<String,HashMap<String, String>> allCommonSheetVariables) {

		//parallelCounter++;
		
		
		allCommonSheetVars = allCommonSheetVariables;
		//System.out.println(parallelCounter);
		int status = -1;     //Declare the status of each method execution to be -1 by default. 
		String row = "";     //To keep track of the currently executing row in the sheet.
		
		appLogs.debug("Called by thread : "  +Thread.currentThread().getName());
		
		
		
		try {
		if(!enable_grid){
			appLogs.info("Launching " + sheetInfo[2] + " browser .....");
			status= launch_browser(sheetInfo);   //This method is used to launch the browser. 
			if (status == 0) {
				return 0;                                      //Stops further execution if the browser launch is failed.
			}	
		}
		else{
			appLogs.info("Launching " + sheetInfo[4] + " browser .....");
			launch_grid_browsers(sheetInfo);
		}
		

		appLogs.info("");
		appLogs.info("< ----------- Starting Execution for sheet : " + sheetInfo[1]
				+ " in File : " + sheetInfo[0] + " ----------- >");

		/* for loop to parse through each row in the sheet*/
		for (int i = 0; i < sheetMatrix.size(); i++) {
			
			row = sheetMatrix.get(i).get(6);   //current running row number taken as a string.
			String action = sheetMatrix.get(i).get(2);  //get the utility method form each row.
			
			/* Print the log message based on the currently executing sheet - Main or Common */
			if (sheetMatrix.get(i).get(0).equals("Main Sheet")) {
				appLogs.info("Currently Executing Row " + row + " : "+ sheetMatrix.get(i).get(1) + " -- "
						+ " in Main Sheet : " + sheetInfo[1] + " -- in File "
						+ " : " + sheetInfo[0]);
			} else {
				appLogs.info("Currently Executing Row " + row + " : "+ sheetMatrix.get(i).get(1) + " -- "
						+ " in Common Sheet : " + sheetMatrix.get(i).get(0)
						+ " in Main Sheet : " + sheetInfo[1] + " in File "
						+ sheetInfo[0]);
			}

			
			switch (action) { //Switch to call a particular utility method starts here.

			case "goto_url":
				status = 0;
				status = goto_url(sheetInfo, sheetMatrix.get(i), fileVariables, row);
				if (status == 0) {
					return 0;
				}
				break;

			case "click":
				status = 0;
				status = click(sheetInfo, sheetMatrix.get(i), fileVariables,
						row );
				if (status == 0) {
					return 0;
				}
				break;

			case "click_wait":
				status = 0;
				status = click_wait(sheetInfo, sheetMatrix.get(i),
						fileVariables, row);
				if (status == 0) {
					return 0;
				}
				break;
				
			case "click_when_clickable":
				status = 0;
				status = click_when_clickable(sheetInfo, sheetMatrix.get(i),
						fileVariables, row);
				if (status == 0) {
					return 0;
				}
				break;	
				
			case "click_if_exists_button_value":
				status = 0;
				status = click_if_exists_button_value(sheetInfo,
						sheetMatrix.get(i), fileVariables, row);
				if (status == 0) {
					return 0;
				}
				break;
				
			case "click_alert_box_ok":
				status = 0;
				status = click_alert_box_ok(sheetInfo, sheetMatrix.get(i),
						fileVariables, row);
				if (status == 0) {
					return 0;
				}
				break;
				
			case "click_alert_box_cancel":
				status = 0;
				status = click_alert_box_cancel(sheetInfo, sheetMatrix.get(i),
						fileVariables, row);
				if (status == 0) {
					return 0;
				}
				break;

		   
			case "enter_text":
				status = 0;
				status = enter_text(sheetInfo, sheetMatrix.get(i),
						fileVariables, row);
				if (status == 0) {
					return 0;
				}
				break;

				
			case "clear_text":
				status = 0;
				status = clear_text(sheetInfo, sheetMatrix.get(i),
						fileVariables, row);
				if (status == 0) {
					return 0;
				}
				break;

				
			case "assert_presence":
				status = 0;
				status = assert_presence(sheetInfo, sheetMatrix.get(i),
						fileVariables, row);
				if (status == 0) {
					return 0;
				}
				break;

				
			case "assert_no_presence":
				status = 0;
				status = assert_no_presence(sheetInfo, sheetMatrix.get(i),
						fileVariables, row);
				if (status == 0) {
					return 0;
				}
				break;

				
			case "assert_text":

				status = 0;
				status = assert_text(sheetInfo, sheetMatrix.get(i),
						fileVariables, row);
				if (status == 0) {
					return 0;
				}
				break;

				
			// not tested yet
			case "mouse_over":
				status = 0;
				status = mouse_over(sheetInfo, sheetMatrix.get(i),
						fileVariables, row);
				if (status == 0) {
					return 0;
				}

				break;

				
			case "click_button":
				break;

				
			case "click_using_button_name":
				break;

				
			case "click_using_button_value":
				break;


			case "click_if_exists":
				break;
				
				
			case "click_button_wait":
				break;

				
			case "select_checkbox_wait":
				break;

				
			case "select_check_box":
				break;

				
			case "select_radion_button":
				break;

				
			case "select_from_dropdown":
				status = 0;
				status = select_from_dropdown(sheetInfo, sheetMatrix.get(i),
						fileVariables, row);
				if (status == 0) {
					return 0;
				}
				break;

				
			case "store_text":
				status = 0;
				status = store_text(sheetInfo, sheetMatrix.get(i),
						fileVariables, row);
				if (status == 0) {
					return 0;
				}
				break;

			
			// Not tested yet
			case "store_value":
				status = 0;
				status = store_value(sheetInfo, sheetMatrix.get(i),
						fileVariables, row);
				if (status == 0) {
					return 0;
				}
				break;

			
			// Not tested yet
			case "store_element_title":
				status = 0;
				status = store_element_title(sheetInfo, sheetMatrix.get(i),
						fileVariables, row);
				if (status == 0) {
					return 0;
				}
				break;

			
			case "upload_file":
				status = 0;
				status = upload_file(sheetInfo, sheetMatrix.get(i),
						fileVariables, row);
				 if (status == 0) {
					return 0;
				}
				break;
			
				
			case "slider_action":
				status = 0;
				status = slider_action(sheetInfo, sheetMatrix.get(i),
						fileVariables, row);
				if (status == 0) {
					return 0;
				}
				break;
				
			case "drag_drop":
				status = 0;
				status = drag_drop(sheetInfo, sheetMatrix.get(i),
						fileVariables, row);
				if (status == 0) {
					return 0;
				}
				break;
				
			case "switch_tab":
				status = 0;
				status = switch_tab(sheetInfo, sheetMatrix.get(i),
						fileVariables, row);
				if (status == 0) {
					return 0;
				}
				break;
		   

			//Not working
			case "assert_query":
				status = 0;
				status = assert_query(sheetInfo, sheetMatrix.get(i),
						fileVariables, row);
				if (status == 0) {
					return 0;
				}
				break;

			case "quit":
				this.quit(sheetInfo, 1);
				break;
			
			default:
				appLogs.error("Execution Failed in Row : " + row + " --> Method " + action + " in sheet " + sheetInfo[1]
						+ " is not valid.");
				/*createResultSet(sheetInfo[0], sheetInfo[1], "FAIL", "Method "+ action + " in step : " + row + " in sheet "
						+ sheetInfo[1] + " is not valid.", "N/A");*/
				createResultSet(sheetInfo, "FAIL", "Method "+ action + " in step : " + row + " in sheet "
						+ sheetInfo[1] + " is not valid.", "N/A");
				quit(sheetInfo, 0);
				return 0;
			} //Switch to call each utility method ends here. 
		}
		
		return 1;
		} 
		catch(Exception e){
			//appLogs.info(parallelCounter);
			appLogs.info("Exception Raised by Thread - " + Thread.currentThread().getName());
			//createResultSet(sheetInfo[0], sheetInfo[1], "FAIL", e.getMessage(),"N/A");
			createResultSet(sheetInfo, "FAIL", e.getMessage(),"N/A");
			return 0;	
			//e.getMessage().toString()
		}
		
		finally{
			//parallelCounter--;
		}
		
	}

	
	/**
	 * Launches the browser to perform the test.
	 * @param browser - the name of the browser to be launched.
	 */
	public int launch_browser(String [] sheetInfo) {
		//WebDriver driver = null;
		
        try{
        	
        	switch(sheetInfo[2]){
        	case "chrome":
        	case "Chrome":
    			System.setProperty("webdriver.chrome.driver",System.getProperty("user.dir")+ "//Input//Drivers//chromedriver.exe");			
    			driver = new ChromeDriver();
        	    break;
        	
        	case "firefox" :
        	case "Firefox" :
        		driver = new FirefoxDriver();
        	    break;
        	
        	case "headless":
        	case "Headless" :
        		caps = new DesiredCapabilities();
        		caps.setJavascriptEnabled(true);	
        		caps.setCapability(PhantomJSDriverService.PHANTOMJS_CLI_ARGS, new String[] {"--web-security=no", "--ignore-ssl-errors=yes"});
        		//caps.setCapability("trustAllSSLCertificates", true);
        		caps.setCapability(PhantomJSDriverService.PHANTOMJS_EXECUTABLE_PATH_PROPERTY, System.getProperty("user.dir")+ "//Input//Drivers//phantomjs.exe");

        		driver = new PhantomJSDriver(caps);
        	break;
        	
        	default :
        		appLogs.error("Execution Failed: Incorrect browser name in the config sheet in file - " + sheetInfo[0]);
        		createResultSet(sheetInfo, "FAIL","Incorrect or Empty browser name in the config sheet in file - " + sheetInfo[0]
        				+ ". Try 'chrome', 'firefox' or 'headless' browser options.", "N/A"); 
        	    return 0;     	
        	
        	}
        	
    		 driver.manage().window().maximize();  //maximize the browser window.	
    		 driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS); // provide implicit wait before throwing NoSuchElementException.
     		 actions = new Actions(driver);  //create object for Actions call for each driver instance 
     		 
     		 return 1;
        }
		catch(Exception e){
	     appLogs.info("Exception in launching the browser by " + Thread.currentThread().getName() + e.getMessage());
	     createResultSet(sheetInfo, "FAIL","Exception while trying to launch the browser.", "N/A"); 
	     return 0;
		}
		
	}
	
	/**
	 * Launches the browsers in the remote machine using selenium Grid
	 * @param SheetInfo: String array containing all the key for each executable sheet.
	 */
	public int launch_grid_browsers(String [] sheetInfo){
		
		String nodeURL = sheetInfo[3];
		String gridBrowser = sheetInfo[4];
		DesiredCapabilities caps = null;
	
	try {
		  switch(gridBrowser){
		  case "chrome":
	  	  case "Chrome":
	  	  case "CHROME":
			  appLogs.info("Inside Chrome Browser for Grid");
			  caps = DesiredCapabilities.chrome();
			  caps.setBrowserName("chrome");
			  caps.setPlatform(Platform.WINDOWS);
			  try {
				driver= new RemoteWebDriver(new URL(nodeURL), caps);
			} catch (MalformedURLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			  break;
			  
	  	 case "firefox" :
		 case "Firefox" :
		 case "FIREFOX" :
			  caps = DesiredCapabilities.firefox();
			  caps.setBrowserName("firefox");
			  caps.setPlatform(Platform.WINDOWS);
			  try {
				driver= new RemoteWebDriver(new URL(nodeURL), caps);
			} catch (MalformedURLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			  break;
			  
		  case "IE11":
			  
			  break;
			  
		  case "IE10":
			  break;
			  
		  case "Edge":
			break;
			
		  default:
			  appLogs.error("Execution Failed: Incorrect browser name in the Master Sheet for the Grid - " + sheetInfo[2]+ "--"  + sheetInfo[3]);
	  		  createResultSet(sheetInfo, "FAIL","Incorrect browser name in the Master Sheet for the Grid - " + sheetInfo[2]+ "--" + sheetInfo[3]
	  				+ ". Try 'chrome', 'firefox' or 'headless' browser options.", "N/A"); 
	  	     return 0;    
			  
		  }
		    
		  driver.manage().window().maximize();  //maximize the browser window.	
		  driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS); // provide implicit wait before throwing NoSuchElementException.
		  actions = new Actions(driver);  //create object for Actions call for each driver instance 
	    	
		  return 1;
	}
	catch(Exception e){
		createResultSet(sheetInfo, "FAIL","Exception was raised while trying to launch"+ sheetInfo[4]+ " the browser for " 
				+sheetInfo[0] +"--" +sheetInfo[1] + "--" + sheetInfo[2]
				+sheetInfo[3], "N/A"); 
		return 0;
	}
	
    }

	/**
	 * Navigates to a particular URL using WebDriver.get(URL) method
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int goto_url(String [] sheetInfo, ArrayList<String> data,HashMap<String, String> variables, String stepNumber) {
		String dataParam = "";
		try {
             /* check if there is any missing parameter in current row  */ 
			if (data.get(5).equals("N/A")) {
				missingArgs(sheetInfo, data, stepNumber);
				return 0;
			}
			
			dataParam = getDataParam(sheetInfo, data, variables, data.get(5));
			driver.get(dataParam);
			return 1;
		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(sheetInfo, data,
					stepNumber, e);
			getScreenshot(sheetInfo, stepNumber, exceptionMessage);
			return 0;
		}

	}

	/**
	 * Enters text in a field using WebDriver.sendKeys() method
	 * 
     * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @return 0 in case of Failure and 1 in case of SuccessfullExecution
	 */
	public int enter_text(String [] sheetInfo, ArrayList<String> data,HashMap<String, String> variables, String stepNumber) {

		String dataParam = "";    //set 'Test Data/Options' value from the list of available variables.
		String locatorValue = ""; // set 'Locator Value' from the list of available variables.
		
		try {
			/* check if there is any missing parameter in current row  */ 
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")|| data.get(5).equals("N/A")) {
				missingArgs(sheetInfo, data, stepNumber);
				return 0;
			}

			Thread.sleep(1000);
			dataParam = getDataParam(sheetInfo, data, variables, data.get(5));
			locatorValue = getDataParam(sheetInfo, data, variables, data.get(4));

			switch (data.get(3)) { // switch starts here.

			case "xpath":
				driver.findElement(By.xpath(locatorValue)).sendKeys(dataParam);
				break;

			case "id":
				driver.findElement(By.id(locatorValue)).sendKeys(dataParam);
				break;

			case "class_name":
				driver.findElement(By.className(locatorValue)).sendKeys(dataParam);
				break;
				
			case "name":
				driver.findElement(By.name(locatorValue)).sendKeys(dataParam);
				break;			
				
			// add more case statements here.

			default:
				incorrectLocatorType(sheetInfo, data, stepNumber);
				return 0;
			}// switch ends here

			return 1;

		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(sheetInfo, data,
					stepNumber, e);
			getScreenshot(sheetInfo, stepNumber, exceptionMessage);
			return 0;
		}
	}
	

	/**
	 * Clears the user input from a web element using WebDriver.clear() method.
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int clear_text(String [] sheetInfo, ArrayList<String> data,HashMap<String, String> variables, String stepNumber) {
		
		String locatorValue = "";
		
		try {
			
			/* check if there is any missing parameter in current row  */ 
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")) {
				missingArgs(sheetInfo, data, stepNumber);
				return 0;
			}

			locatorValue = getDataParam(sheetInfo, data, variables, data.get(4));
			Thread.sleep(1000);
			switch (data.get(3)) { // switch starts here.

			case "xpath":
				driver.findElement(By.xpath(locatorValue)).clear();
				;
				break;

			case "id":
				driver.findElement(By.id(locatorValue)).clear();
				break;

			case "class_name":
				driver.findElement(By.className(locatorValue)).clear();
				break;
				
			case "name":
				driver.findElement(By.name(locatorValue)).clear();
				break;		

			// add more case statements here.

			default:
				incorrectLocatorType(sheetInfo, data, stepNumber);
				return 0;
			}// switch ends here

			return 1;

		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(sheetInfo, data,	stepNumber, e);
			getScreenshot(sheetInfo, stepNumber, exceptionMessage);
			return 0;
		}
	}

	/**
	 * Clicks on a particular element using WebDriver.click() method.
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int click(String [] sheetInfo, ArrayList<String> data,HashMap<String, String> variables, String stepNumber) {

		String locatorValue = "";
		
		try {
			
			/* check if required parameters for this action are missing or not */ 
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")) {
				missingArgs(sheetInfo, data, stepNumber);
				return 0;
			}

			locatorValue = getDataParam(sheetInfo, data, variables, data.get(4));
			Thread.sleep(1000);
			switch (data.get(3)) { // switch starts here.

			case "xpath":
				driver.findElement(By.xpath(locatorValue)).click();
				break;

			case "id":
				driver.findElement(By.id(locatorValue)).click();
				break;

			case "class_name":
				driver.findElement(By.className(locatorValue)).click();
				break;
				
			case "name":
				driver.findElement(By.name(locatorValue)).click();
				break;	
						
			case "linkText":
				driver.findElement(By.linkText(locatorValue)).click();
				break;	
				
			case "partialLinkText":
				driver.findElement(By.partialLinkText(locatorValue)).click();
				break;
				
			case "title":
				driver.findElement(By.xpath("//*[@title = '" + locatorValue + "']")).click();
				break;	
						
			case "value":
				driver.findElement(By.xpath("//*[@value = '" + locatorValue + "']")).click();
				break;	
				

			// add more case statements here.

			default:
				incorrectLocatorType(sheetInfo, data, stepNumber);
				return 0;
			}// switch ends here

			return 1;
			
		} catch (Exception e) {
			String exceptionMessage = exceptionalCondition(sheetInfo, data,
					stepNumber, e);
			getScreenshot(sheetInfo, stepNumber, exceptionMessage);
			return 0;
		}

	}

	/**
	 * Clicks on an element using its value attribute, only if that element is present on the page.
	 * 
     * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int click_if_exists_button_value(String [] sheetInfo,ArrayList<String> data, HashMap<String, String> variables,String stepNumber) {

		String locatorValue = "";
		try {

			/* check if required parameters for this action are missing or not */ 
			if (data.get(4).equals("N/A")) {
				missingArgs(sheetInfo, data, stepNumber);
				return 0;
			}

			locatorValue = getDataParam(sheetInfo, data, variables, data.get(4));
			
			//Check if the element based on the provided value attribute is present or not. Click if present.
			if (!driver.findElements(
					By.xpath("//*[@value = '" + locatorValue + "']")).isEmpty()) {
				((WebElement) driver.findElement(By.xpath("//*[@value = '"
						+ locatorValue + "']"))).click();
			}

			else {
				return 1;
			}

			return 1;
		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(sheetInfo, data,
					stepNumber, e);
			getScreenshot(sheetInfo, stepNumber, exceptionMessage);
			return 0;
		}
	}

	/**
	 * Wait a particular element to appear on the page then click it. 
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */

	public int click_wait(String [] sheetInfo, ArrayList<String> data, HashMap<String, String> variables, String stepNumber) {
		
		wait = new WebDriverWait(driver, 30);  //Creating a new WebDriverWait and passing the WebDriver and TimeOut
		String locatorValue = " ";
		
		try {
			
			/* check if required parameters for this action are missing or not */ 
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")) {
				missingArgs(sheetInfo, data, stepNumber);
				return 0;
			}

			locatorValue = getDataParam(sheetInfo, data, variables, data.get(4));

			switch (data.get(3)) { // switch starts here.

			case "xpath":
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(locatorValue)));
				actions.moveToElement(driver.findElement(By.xpath(locatorValue))).click().perform();
				break;

			case "id":
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(locatorValue)));
				actions.moveToElement(driver.findElement(By.id(locatorValue))).click().perform();
				break;

			case "class_name":
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.className(locatorValue)));
				actions.moveToElement(driver.findElement(By.className(locatorValue))).click().perform();
				break;
				
			case "name":
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.className(locatorValue)));
				actions.moveToElement(driver.findElement(By.name(locatorValue))).click().perform();
				break;	
						
			case "linkText":
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.className(locatorValue)));
				actions.moveToElement(driver.findElement(By.linkText(locatorValue))).click().perform();
				break;	
				
			case "partialLinkText":
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.className(locatorValue)));
				actions.moveToElement(driver.findElement(By.partialLinkText(locatorValue))).click().perform();
				break;
				
			case "title":
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.className(locatorValue)));
				actions.moveToElement(driver.findElement(By.xpath("//*[@title = '" + locatorValue + "']"))).click().perform();
				break;	
						
			case "value":
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.className(locatorValue)));
				actions.moveToElement(driver.findElement(By.xpath("//*[@value = '" + locatorValue + "']"))).click().perform();
				break;	

			// add more case statements here.

			default:
				incorrectLocatorType(sheetInfo, data, stepNumber);
				return 0;
			}// switch ends here

			return 1;
		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(sheetInfo, data,
					stepNumber, e);
			getScreenshot(sheetInfo, stepNumber, exceptionMessage);
			return 0;
		}
	}
	
	/**
	 * Click on a particular web element only when its clickable.
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @param currentSheet - Name of the current executing Sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	
	public int click_when_clickable(String[] sheetInfo, ArrayList<String> data, HashMap<String, String> variables, String stepNumber){
		
		wait = new WebDriverWait(driver, 20);  //Creating a new WebDriverWait and passing the WebDriver and TimeOut
	
		String locatorValue = "";
		
		try {
			
			/* check if required parameters for this action are missing or not */ 
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")) {
				missingArgs(sheetInfo, data, stepNumber);
				return 0;
			}

			locatorValue = getDataParam(sheetInfo, data, variables, data.get(4));
			Thread.sleep(1000);
			switch (data.get(3)) { // switch starts here.

			case "xpath":
				wait.until(ExpectedConditions.elementToBeClickable(By.xpath(locatorValue)));
				actions.moveToElement(driver.findElement(By.xpath(locatorValue))).click().perform();
				break;

			case "id":
				wait.until(ExpectedConditions.elementToBeClickable(By.id(locatorValue)));
				actions.moveToElement(driver.findElement(By.id(locatorValue))).click().perform();
				break;

			case "class_name":
				wait.until(ExpectedConditions.elementToBeClickable(By.className(locatorValue)));
				actions.moveToElement(driver.findElement(By.className(locatorValue))).click().perform();
				break;
				
			case "name":
				wait.until(ExpectedConditions.elementToBeClickable(By.className(locatorValue)));
				actions.moveToElement(driver.findElement(By.name(locatorValue))).click().perform();
				break;	
						
			case "linkText":
				wait.until(ExpectedConditions.elementToBeClickable(By.className(locatorValue)));
				actions.moveToElement(driver.findElement(By.linkText(locatorValue))).click().perform();
				break;	
				
			case "partialLinkText":
				wait.until(ExpectedConditions.elementToBeClickable(By.className(locatorValue)));
				actions.moveToElement(driver.findElement(By.partialLinkText(locatorValue))).click().perform();
				break;
				
			case "title":
				wait.until(ExpectedConditions.elementToBeClickable(By.className(locatorValue)));
				actions.moveToElement(driver.findElement(By.xpath("//*[@title = '" + locatorValue + "']"))).click().perform();
				break;	
						
			case "value":
				wait.until(ExpectedConditions.elementToBeClickable(By.className(locatorValue)));
				actions.moveToElement(driver.findElement(By.xpath("//*[@value = '" + locatorValue + "']"))).click().perform();
				break;	


			// add more case statements here.

			default:
				incorrectLocatorType(sheetInfo, data, stepNumber);
				return 0;
			}// switch ends here

			return 1;
			
		} catch (Exception e) {
			String exceptionMessage = exceptionalCondition(sheetInfo, data,
					stepNumber, e);
			getScreenshot(sheetInfo, stepNumber, exceptionMessage);
			return 0;
		}
				
	}

	/**
	 * Assert that a particular element is displayed on the page based on its text value.
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int assert_text(String [] sheetInfo, ArrayList<String> data, HashMap<String, String> variables, String stepNumber) {

		String actualValue = "";
		String dataParam = "";
		String locatorValue = "";
		
		try {
			
			/* check if required parameters for this action are missing or not */ 
			
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")|| data.get(5).equals("N/A")) {
				missingArgs(sheetInfo, data, stepNumber);
				return 0;
			}

			locatorValue =getDataParam(sheetInfo, data, variables, data.get(4));
			dataParam = getDataParam(sheetInfo, data, variables, data.get(5));

			switch (data.get(3)) { // switch starts here.

			case "xpath":
				actualValue = driver.findElement(By.xpath(locatorValue)).getText();
				break;

			case "id":
				actualValue = driver.findElement(By.id(locatorValue)).getText();
				break;

			case "class_name":
				actualValue = driver.findElement(By.className(locatorValue)).getText();
				break;
				
			case "name":
				actualValue = driver.findElement(By.name(locatorValue)).getText();
				break;	
						
			case "linkText":
				actualValue = driver.findElement(By.linkText(locatorValue)).getText();
				break;	
				
			case "partialLinkText":
				actualValue = driver.findElement(By.partialLinkText(locatorValue)).getText();
				break;
				
			case "title":
				actualValue = driver.findElement(By.xpath("//*[@title = '" + locatorValue + "']")).getText();
				break;	
						
			case "value":
				actualValue = driver.findElement(By.xpath("//*[@value = '" + locatorValue + "']")).getText();
				break;	

			// add more case statements here.

			default:
				incorrectLocatorType(sheetInfo, data, stepNumber);
				return 0;
			}// switch ends here

			if (!(actualValue.equals(dataParam))) {
				appLogs.error("Execution Failed: Actual text on platform does not match with Argument Provided for Action "
						+ data.get(2)+ " in Step number "+ stepNumber+ " of Sheet " + sheetInfo[1]);
				appLogs.info("Actual Value :" + actualValue + "-----" + " Exceted Value : " + dataParam);

				getScreenshot(sheetInfo, stepNumber,"Execution Failed at Step " + stepNumber+ " due to data mismatch - "
								+ " Actual Value :" + actualValue + " ----- "+ " Expected Value : " + dataParam);
				return 0;

			}
			return 1;
		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(sheetInfo, data,stepNumber, e);
			getScreenshot(sheetInfo, stepNumber, exceptionMessage);
			return 0;
		}
	}

	/**
	 * assert the presence of an element on the page based on its locator value.
	 * 
     * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int assert_presence(String [] sheetInfo, ArrayList<String> data, HashMap<String, String> variables, String stepNumber) {
		
		String locatorValue = "";
		
		try {
			/* check if required parameters for this action are missing or not */ 
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")) {
				missingArgs(sheetInfo, data, stepNumber);
				return 0;
			}

			locatorValue = getDataParam(sheetInfo, data, variables, data.get(4));

			switch (data.get(3)) { // switch starts here.

			case "xpath":
				driver.findElement(By.xpath(locatorValue));
				break;

			case "id":
				driver.findElement(By.id(locatorValue));
				break;

			case "class_name":
				driver.findElement(By.className(locatorValue));
				break;
				
			case "name":
				driver.findElement(By.name(locatorValue));
				break;	
						
			case "linkText":
				driver.findElement(By.linkText(locatorValue));
				break;	
				
			case "partialLinkText":
				driver.findElement(By.partialLinkText(locatorValue));
				break;
				
			case "title":
				driver.findElement(By.xpath("//*[@title = '" + locatorValue + "']"));
				break;	
						
			case "value":
				driver.findElement(By.xpath("//*[@value = '" + locatorValue + "']"));
				break;	

			// add more case statements here.

			default:
				incorrectLocatorType(sheetInfo, data, stepNumber);
				return 0;
			}// switch ends here

			return 1;
		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(sheetInfo, data,
					stepNumber, e);
			getScreenshot(sheetInfo, stepNumber, exceptionMessage);
			return 0;
		}
	}

	/**
	 * Assert that an element is not present on the page based on its locator value 
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */

	public int assert_no_presence(String[] sheetInfo, ArrayList<String> data,HashMap<String, String> variables, String stepNumber) {
		
		String locatorValue = "";
		List numberOfElements = new ArrayList();
		
		try {
			/* check if required parameters for this action are missing or not */ 
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")) {
				missingArgs(sheetInfo, data, stepNumber);
				return 0;
			}

			locatorValue = getDataParam(sheetInfo, data, variables, data.get(4));
			
			switch (data.get(3)) { // switch starts here.

			case "xpath":
				numberOfElements = driver.findElements(By.xpath(locatorValue));
				break;

			case "id":
				numberOfElements = driver.findElements(By.id(locatorValue));
				break;

			case "class_name":
				numberOfElements = driver.findElements(By.className(locatorValue));
				break;
				
			case "name":
				numberOfElements = driver.findElements(By.name(locatorValue));
				break;	
						
			case "linkText":
				numberOfElements = driver.findElements(By.linkText(locatorValue));
				break;	
				
			case "partialLinkText":
				numberOfElements = driver.findElements(By.partialLinkText(locatorValue));
				break;
				
			case "title":
				numberOfElements = driver.findElements(By.xpath("//*[@title = '" + locatorValue + "']"));
				break;	
						
			case "value":
				numberOfElements = driver.findElements(By.xpath("//*[@value = '" + locatorValue + "']"));
				break;	

			// add more case statements here.

			default:
				incorrectLocatorType(sheetInfo, data, stepNumber);
				return 0;
			}// switch ends here

			if (numberOfElements.size() > 0) {

				appLogs.error("Execution Failed :  For Action " + data.get(2)
						+ " in Step number " + stepNumber + " of Sheet "
						+ sheetInfo[1]);
				getScreenshot(sheetInfo, stepNumber, "Assertion Failed :  For Action " + data.get(2)
						+ " in Step number " + stepNumber + " of Sheet " + sheetInfo[1] );
				return 0;
			} 
			
			else 
				return 1;

		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(sheetInfo, data,stepNumber, e);
			getScreenshot(sheetInfo, stepNumber, exceptionMessage);
			return 0;
		}
	}

	/**
	 * Verify the browser alert box on the page and click on 'OK' button. 
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int click_alert_box_ok(String[] sheetInfo, ArrayList<String> data, HashMap<String, String> variables, String stepNumber) {

		Alert alert = driver.switchTo().alert();

		String dataParam = "";
		String locatorValue = "";

		try {
			/* check if required parameters for this action are missing or not */ 
			if (data.get(5).equals("N/A")) {
				missingArgs(sheetInfo, data, stepNumber);
				return 0;
			}

			dataParam = getDataParam(sheetInfo, data, variables, data.get(5));

			/*Assert the text in the Alert pop up */
			if (!(alert.getText().equals(dataParam))) {
				appLogs.error("Execution Failed: Alert Text does not match with Argument Provided for Action "
						+ data.get(2)+ " in Step number "+ stepNumber+ " of Sheet " + sheetInfo[1]);
				appLogs.info("Actual Value :" + alert.getText() + "-----"+ " Exceted Value : " + dataParam);

				getScreenshot(sheetInfo, stepNumber,"Execution Failed at Step " + stepNumber
								+ " due to in correct alert pop up message - "+ " Actual Value :" + alert.getText()
								+ " ----- " + " Expected Value : " + dataParam);
				return 0;
			}

			alert.accept(); // Clicks on OK button in the Alert window.

			return 1;

		} 
		
		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(sheetInfo, data,stepNumber, e);
			getScreenshot(sheetInfo, stepNumber, exceptionMessage);
			return 0;
		}
	}
	
	/**
	 * Verify the browser alert box on the page and click on 'OK' button. 
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int click_alert_box_cancel(String [] sheetInfo, ArrayList<String> data, HashMap<String, String> variables, String stepNumber) {

		Alert alert = driver.switchTo().alert();

		String dataParam = "";
		String locatorValue = "";

		try {
			/* check if required parameters for this action are missing or not */ 
			if (data.get(5).equals("N/A")) {
				missingArgs(sheetInfo, data, stepNumber);
				return 0;
			}

			dataParam = getDataParam(sheetInfo, data, variables, data.get(5));

			/*Assert the text in the Alert pop up */
			if (!(alert.getText().equals(dataParam))) {
				appLogs.error("Execution Failed: Alert Text does not match with Argument Provided for Action "
						+ data.get(2)+ " in Step number "+ stepNumber+ " of Sheet " + sheetInfo[1]);
				appLogs.info("Actual Value :" + alert.getText() + "-----"+ " Exceted Value : " + dataParam);

				getScreenshot(sheetInfo, stepNumber,"Execution Failed at Step " + stepNumber
								+ " due to in correct alert pop up message - "+ " Actual Value :" + alert.getText()
								+ " ----- " + " Expected Value : " + dataParam);
				return 0;
			}

			alert.dismiss(); // Clicks on dismiss button in the Alert window.

			return 1;

		} catch (Exception e) {
			String exceptionMessage = exceptionalCondition(sheetInfo, data,stepNumber, e);
			getScreenshot(sheetInfo, stepNumber, exceptionMessage);
			return 0;
		}
	}


	/**
	 * Select a element from the drop down (to be used only with the web elements that have 
	 *  'Select' tag)
	 *  
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int select_from_dropdown(String [] sheetInfo, ArrayList<String> data,HashMap<String, String> variables, String stepNumber) {

		String dataParam = "";
		String locatorValue = "";
		try {
			/* check if required parameters for this action are missing or not */ 
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")|| data.get(5).equals("N/A")) {
				missingArgs(sheetInfo, data, stepNumber);
				return 0;
			}
			
			dataParam = getDataParam(sheetInfo, data, variables, data.get(5));
			locatorValue = getDataParam(sheetInfo, data, variables, data.get(4));

			switch (data.get(3)) { // switch starts here.

			case "xpath":
				new Select(driver.findElement(By.xpath(locatorValue))).selectByVisibleText(dataParam);
				break;

			case "id":
				new Select(driver.findElement(By.id(locatorValue))).selectByVisibleText(dataParam);
				break;

			case "class_name":
				new Select(driver.findElement(By.className(locatorValue))).selectByVisibleText(dataParam);
				break;

			// add more case statements here.

			default:
				incorrectLocatorType(sheetInfo, data, stepNumber);
				return 0;
			}// switch ends here

			return 1;
		} 
		
		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(sheetInfo, data,stepNumber, e);
			getScreenshot(sheetInfo, stepNumber, exceptionMessage);
			return 0;
		}
	}

	/**
	 * Hover the mouse over a web element.
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int mouse_over(String[] sheetInfo, ArrayList<String> data,HashMap<String, String> variables, String stepNumber) {
		
		String locatorValue = "";
		
		try {
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")) {
				missingArgs(sheetInfo, data, stepNumber);
				return 0;
			}

			locatorValue = getDataParam(sheetInfo, data, variables, data.get(4));
			switch (data.get(3)) { // switch starts here.

			case "xpath":
				actions.moveToElement(driver.findElement(By.xpath(locatorValue))).build().perform();
				break;

			case "id":
				actions.moveToElement(driver.findElement(By.id(locatorValue))).build().perform();
				break;

			case "class_name":
				actions.moveToElement(driver.findElement(By.className(locatorValue))).build().perform();
				break;

			// add more case statements here.

			default:
				incorrectLocatorType(sheetInfo, data, stepNumber);
				return 0;
			}// switch ends here

			return 1;

		} catch (Exception e) {
			String exceptionMessage = exceptionalCondition(sheetInfo, data,stepNumber, e);
			getScreenshot(sheetInfo,  stepNumber, exceptionMessage);
			return 0;
		}
	}

	
	/**
	 * Store the text of an element in a variable and add this variable to 'variables' hashmap for further usage.
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int store_text(String[] sheetInfo, ArrayList<String> data,HashMap<String, String> variables, String stepNumber) {
		
		String key = null;
		String value = null;
		String locatorValue = "";
		
		try {
			
			/* check if required parameters for this action are missing or not */ 
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")|| data.get(5).equals("N/A")) {
				missingArgs(sheetInfo, data, stepNumber);
				return 0;
			}

			locatorValue = getDataParam(sheetInfo, data, variables, data.get(4));
			
			if(data.get(5).toString().substring(0, 5).equals("var <")){
				key = data.get(5).toString().substring(3).trim();
			}
			else{
				incorrectVariableDeclaration(sheetInfo, data, stepNumber);
				return 0;
			}

			switch (data.get(3)) { // switch starts here.

			case "xpath":
				value = driver.findElement(By.xpath(locatorValue)).getText();
				break;

			case "id":
				value = driver.findElement(By.id(locatorValue)).getText();
				break;

			case "class_name":
				value = driver.findElement(By.className(locatorValue)).getText();
				break;
				
			case "name":
				value = driver.findElement(By.name(locatorValue)).getText();
				break;	
						
			case "linkText":
				value = driver.findElement(By.linkText(locatorValue)).getText();
				break;	
				
			case "partialLinkText":
				value = driver.findElement(By.partialLinkText(locatorValue)).getText();
				break;
				
			case "title":
				value = driver.findElement(By.xpath("//*[@title = '" + locatorValue + "']")).getText();
				break;	
						
			case "value":
				value = driver.findElement(By.xpath("//*[@value = '" + locatorValue + "']")).getText();
				break;	


			// add more case statements here.

			default:
				incorrectLocatorType(sheetInfo, data, stepNumber);
				return 0;
			}// switch ends here

			variables.put(key, value);

			return 1;
		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(sheetInfo, data,stepNumber, e);
			getScreenshot(sheetInfo, stepNumber, exceptionMessage);
			return 0;
		}
	}

	
	/**
	 * Store the value of an element in a variable and add it to the 'Variables' hashmap for further usage.
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
     * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int store_value(String[] sheetInfo, ArrayList<String> data,HashMap<String, String> variables, String stepNumber) {
		
		String key = null;
		String value = null;
		String locatorValue = "";
		
		try {
            
			/* check if required parameters for this action are missing or not */ 
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")|| data.get(5).equals("N/A")) {
				missingArgs(sheetInfo, data, stepNumber);
				return 0;
			}

			locatorValue = getDataParam(sheetInfo, data, variables, data.get(4));
			if(data.get(5).toString().substring(0, 5).equals("var <")){
				key = data.get(5).toString().substring(3).trim();
			}
			else{
				incorrectVariableDeclaration(sheetInfo, data, stepNumber);
				return 0;
			}

			switch (data.get(3)) { // switch starts here.

			case "xpath":
				value = driver.findElement(By.xpath(locatorValue)).getAttribute("value");
				break;

			case "id":
				value = driver.findElement(By.id(locatorValue)).getAttribute("value");
				break;

			case "class_name":
				value = driver.findElement(By.className(locatorValue)).getAttribute("value");
				break;
				
			case "name":
				value = driver.findElement(By.name(locatorValue)).getAttribute("value");
				break;	
						
			case "linkText":
				value = driver.findElement(By.linkText(locatorValue)).getAttribute("value");
				break;	
				
			case "partialLinkText":
				value = driver.findElement(By.partialLinkText(locatorValue)).getAttribute("value");
				break;
				
			case "title":
				value = driver.findElement(By.xpath("//*[@title = '" + locatorValue + "']")).getAttribute("value");
				break;	
						
			case "value":
				value = driver.findElement(By.xpath("//*[@value = '" + locatorValue + "']")).getAttribute("value");
				break;	

			// add more case statements here.

			default:
				incorrectLocatorType(sheetInfo, data, stepNumber);
				return 0;
			}// switch ends here

			variables.put(key, value);

			return 1;
		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(sheetInfo, data, stepNumber, e);
			getScreenshot(sheetInfo, stepNumber, exceptionMessage);
			return 0;
		}
	}

	
	/**
	 * Store the title attribute of a web Element in a variable and store this in 'Variables'  hashMap for further usage
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int store_element_title(String[] sheetInfo, ArrayList<String> data, HashMap<String, String> variables, String stepNumber) {
	
		String key = null;
		String value = null;
		String locatorValue = "";
		
		try {
 
			/* check if required parameters for this action are missing or not */ 
			if (data.get(3).equals("N/A") || data.get(4).equals("N/A")|| data.get(5).equals("N/A")) {
				missingArgs(sheetInfo, data, stepNumber);
				return 0;
			}

			locatorValue = getDataParam(sheetInfo, data, variables, data.get(4));
			if(data.get(5).toString().substring(0, 5).equals("var <")){
				key = data.get(5).toString().substring(3).trim();
			}
			else{
				incorrectVariableDeclaration(sheetInfo, data, stepNumber);
				return 0;
			}
			

			switch (data.get(3)) { // switch starts here.

			case "xpath":
				value = driver.findElement(By.xpath(locatorValue)).getAttribute("title");
				break;

			case "id":
				value = driver.findElement(By.id(locatorValue)).getAttribute("title");
				break;

			case "class_name":
				value = driver.findElement(By.className(locatorValue)).getAttribute("title");
				break;
				
			case "name":
				value = driver.findElement(By.name(locatorValue)).getAttribute("title");
				break;	
						
			case "linkText":
				value = driver.findElement(By.linkText(locatorValue)).getAttribute("title");
				break;	
				
			case "partialLinkText":
				value = driver.findElement(By.partialLinkText(locatorValue)).getAttribute("title");
				break;
				
			case "title":
				value = driver.findElement(By.xpath("//*[@title = '" + locatorValue + "']")).getAttribute("title");
				break;	
						
			case "value":
				value = driver.findElement(By.xpath("//*[@value = '" + locatorValue + "']")).getAttribute("title");
				break;	

			// add more case statements here.

			default:
				incorrectLocatorType(sheetInfo, data, stepNumber);
				return 0;
			}// switch ends here

			variables.put(key, value);

			return 1;
		}

		catch (Exception e) {
			String exceptionMessage = exceptionalCondition(sheetInfo, data,stepNumber, e);
			getScreenshot(sheetInfo, stepNumber, exceptionMessage);
			return 0;
		}
	}

	/**
	 * Upload a file 
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.

	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */

	public int upload_file(String[] sheetInfo, ArrayList<String> data,	HashMap<String, String> variables, String stepNumber) {

		String dataParam = "";
		String locatorValue = " ";
		
		/* check if required parameters for this action are missing or not */ 
		if (data.get(3).equals("N/A") || data.get(4).equals("N/A")|| data.get(5).equals("N/A")) {
			missingArgs(sheetInfo, data, stepNumber);
			return 0;
			
		}

		dataParam = getDataParam(sheetInfo, data, variables, data.get(5));
		locatorValue = getDataParam(sheetInfo, data, variables, data.get(4));
		JavascriptExecutor js = (JavascriptExecutor) driver;
		
		try{
			
			switch (data.get(3)) { // switch starts here.

			case "xpath":
				js.executeScript("document.getElementByXpath('"+locatorValue+"').style.display= 'block';");
				driver.findElement(By.xpath(locatorValue)).sendKeys(dataParam);
				//js.executeScript("document.getElementByXpath('"+locatorValue+"').style.display = 'none';");

				break;

			case "id":
				js.executeScript("document.getElementById('"+locatorValue+"').style.display= 'block';");
				driver.findElement(By.id(locatorValue)).sendKeys(dataParam);
				//js.executeScript("document.getElementById('"+locatorValue+"').style.display = 'none';");

				break;

			case "class_name":
				js.executeScript("document.getElementByClassName('"+locatorValue+"').style.display= 'block';");
				driver.findElement(By.className(locatorValue)).sendKeys(dataParam);
				break;

				
			// add more case statements here.

			default:
				incorrectLocatorType(sheetInfo, data, stepNumber);
				return 0;
				
			}// switch ends here
		
			return 1;
			
		}
		catch(Exception e){
			String exceptionMessage = exceptionalCondition(sheetInfo, data,stepNumber, e);
			getScreenshot(sheetInfo, stepNumber, exceptionMessage);
			return 0;				
		}
		
	}
	
	
	/**
	 * Perform Slider action.
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.

	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	
	public int slider_action(String [] sheetInfo, ArrayList<String> data,HashMap<String, String> variables, String stepNumber){
		
		WebElement toSlide = null; //WebElement that is to be slided.
		String xoffset = ""; //horizontal offset value 
		String yoffset = ""; //vertical offset value
		String locatorValue = " "; //xpath, id, etc., locator value of the WebElement
		
		/* check if required parameters for this action are missing or not */ 
		if (data.get(3).equals("N/A") || data.get(4).equals("N/A")|| data.get(5).equals("N/A")) {
			missingArgs(sheetInfo, data, stepNumber);
			return 0;
			
		}
		
		if(data.get(5).split(",").length!=2){
			appLogs.error("Execution Failed: Missing Arugments for the Action " + data.get(2) 
				    + " in Step number " + stepNumber + " of Sheet "+ sheetInfo[1]);
		
		   createResultSet(sheetInfo, "FAIL","Incorrect offset Format provided for the action " + data.get(2)
						+ " in Step number " + stepNumber + " of Sheet "+ sheetInfo[1], "N/A");
		
		   quit(sheetInfo, 0); //Quit the browser due to Execution Failure
		   
		   return 0;
		}
		
		locatorValue = getDataParam(sheetInfo, data, variables, data.get(4));
		xoffset= getDataParam(sheetInfo, data, variables, data.get(5).split(",")[0].trim());
		yoffset= getDataParam(sheetInfo, data, variables, data.get(5).split(",")[1].trim());
		
		try{
			
		
		switch (data.get(3)) { // switch starts here.

		case "xpath":
			toSlide= driver.findElement(By.xpath(locatorValue));
			break;

		case "id":
			toSlide = driver.findElement(By.id(locatorValue));
			break;

		case "class_name":
			toSlide = driver.findElement(By.className(locatorValue));
			break;
			
		case "name":
			toSlide = driver.findElement(By.name(locatorValue));
			break;	
					
		case "linkText":
			toSlide = driver.findElement(By.linkText(locatorValue));
			break;	
			
		case "partialLinkText":
			toSlide = driver.findElement(By.partialLinkText(locatorValue));
			break;
			
		case "title":
			toSlide = driver.findElement(By.xpath("//*[@title = '" + locatorValue + "']"));
			break;	
					
		case "value":
			toSlide = driver.findElement(By.xpath("//*[@value = '" + locatorValue + "']"));
			break;	

		// add more case statements here.

		default:
			incorrectLocatorType(sheetInfo, data, stepNumber);
			return 0;
		}// switch ends here
		
		
		actions.dragAndDropBy(toSlide, Integer.parseInt(xoffset), Integer.parseInt(yoffset)).build().perform(); //Perform Slide functionality over here.
		
		return 1;
		
		
		}
		catch(Exception e){
			String exceptionMessage = exceptionalCondition(sheetInfo, data,stepNumber, e);
			getScreenshot(sheetInfo, stepNumber, exceptionMessage);
			return 0;		
		}
	}
	
	
	/**
	 * Drags one element on to another.
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.

	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	
	
	public int drag_drop(String [] sheetInfo, ArrayList<String> data,HashMap<String, String> variables, String stepNumber){
		
		WebElement source = null;   //WebElement that is to be dragged.
		WebElement target = null; //WebElement that is to dragged on to.
		String sourceLocatorType = ""; //Locator Type of the element that is to be dragged.
		String targetLocatorType = ""; //Locator Type of the element that is to be dragged.
		String sourceLocatorValue = "";   //Locator Value of the element that is to be dragged.
		String targetLocatorValue = "";   //Locator value of the element that is to be dragged on to.
		
		/* check if required parameters for this action are missing or not */ 
		if (data.get(3).equals("N/A") || data.get(3).split(";;").length!=2 || 
				data.get(4).equals("N/A") || data.get(4).split(";;").length!=2) {
			missingArgs(sheetInfo, data, stepNumber);
			return 0;
			
		}
		
		/*get the locator value both the elements*/
		sourceLocatorType = data.get(3).split(";;")[0];
		targetLocatorType = data.get(3).split(";;")[1];
		
		/*Get the Locator Value for both the elements */
		sourceLocatorValue = getDataParam(sheetInfo, data, variables, data.get(4).split(";;")[0]);
		targetLocatorValue = getDataParam(sheetInfo, data, variables, data.get(4).split(";;")[1]);
		
      try{
			
			switch (sourceLocatorType) { // switch starts here.

			case "xpath":
				source = driver.findElement(By.xpath(sourceLocatorValue));
				break;

			case "id":
				source = driver.findElement(By.id(sourceLocatorValue));
				break;

			case "class_name":
				source = driver.findElement(By.className(sourceLocatorValue));
				break;
				
			case "name":
				source = driver.findElement(By.name(sourceLocatorValue));
				break;	
						
			case "linkText":
				source =  driver.findElement(By.linkText(sourceLocatorValue));
				 break;
				
			case "partialLinkText":
				source = driver.findElement(By.partialLinkText(sourceLocatorValue));
				break;
				
			case "title":
				source = driver.findElement(By.xpath("//*[@title = '" + sourceLocatorValue + "']"));
				break;	
						
			case "value":
				source = driver.findElement(By.xpath("//*[@value = '" + sourceLocatorValue + "']"));
				break;
            
				
			// add more case statements here.

			default:
				incorrectLocatorType(sheetInfo, data, stepNumber);
				return 0;
				
			}// switch ends here
			
			
			switch (targetLocatorType) { // switch starts here.

			case "xpath":
				target = driver.findElement(By.xpath(targetLocatorValue));
				break;

			case "id":
				target = driver.findElement(By.id(targetLocatorValue));
				break;

			case "class_name":
				target = driver.findElement(By.className(targetLocatorValue));
				break;
				
			case "name":
				target = driver.findElement(By.name(targetLocatorValue));
				break;	
						
			case "linkText":
				target =  driver.findElement(By.linkText(targetLocatorValue));
				 break;
				
			case "partialLinkText":
				target = driver.findElement(By.partialLinkText(targetLocatorValue));
				break;
				
			case "title":
				target = driver.findElement(By.xpath("//*[@title = '" + targetLocatorValue + "']"));
				break;	
						
			case "value":
				target = driver.findElement(By.xpath("//*[@value = '" + targetLocatorValue + "']"));
				break;
            
				
			// add more case statements here.

			default:
				incorrectLocatorType(sheetInfo, data, stepNumber);
				return 0;
				
			}// switch ends here
		
			appLogs.info("Source Element: "  + source.toString() + " -- "+ sourceLocatorType + " --" + sourceLocatorValue );
			appLogs.info("Target Element: " + target.toString() + " -- " + targetLocatorType + " --" + targetLocatorValue );
			
			 //Both the below actions are not working. 
			//actions.dragAndDrop(source, target).build().perform();   //Perform the drag operation.
			//actions.clickAndHold(source).moveToElement(target).release(target).build().perform(); //Perform the drag operation.
			
			actions.clickAndHold(source).perform();
			Thread.sleep(1000);
			actions.moveToElement(target).perform();
			Thread.sleep(1000);
			actions.release(target).perform();
			 
			return 1;
			
		}
		catch(Exception e){
			String exceptionMessage = exceptionalCondition(sheetInfo, data,stepNumber, e);
			getScreenshot(sheetInfo, stepNumber, exceptionMessage);
			return 0;				
		}		
		
	}
	
	

	/**
	 * Use to switch to a tab which does not have the focus.
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	
	public int switch_tab(String [] sheetInfo, ArrayList<String> data,HashMap<String, String> variables, String stepNumber){
		
	 try{
		 String parentHandle = driver.getWindowHandle(); //get the Handle of the currently focused tab.
			Set<String>handles = driver.getWindowHandles();
			
			for(String handle : handles){
				if(!handle.equalsIgnoreCase(parentHandle)){
					driver.switchTo().window(handle);
				}
			}
			
			return 1;
	 }
	 catch(Exception e){
		 String exceptionMessage = exceptionalCondition(sheetInfo, data,stepNumber, e);
		 getScreenshot(sheetInfo, stepNumber, exceptionMessage);
		 return 0;	
	 }	
	}
	

	/**
	 * 
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - ArrayList of all the relevant cell data in the currently executing Row.
	 * @param variables - All the variables (specified in config sheet and during run time) for the file
	 * @param stepNumber - Currently executing Row in the sheet.
	 * @return 0 in case of Failure and 1 in case of Successfully Execution
	 */
	public int assert_query(String [] sheetInfo, ArrayList<String> data,HashMap<String, String> variables, String stepNumber) {

		String URL = "jdbc:mysql://mysql-proxy.brightedge.com:13700/optiweber2";
		String userName = "readyonly";
		String password = "granted";
		String query = "";
		String dataParam = "";
		String locatorValue = "";

		try {
			/* check if required parameters for this action are missing or not */ 
			if (data.get(4).equals("N/A") || data.get(5).equals("N/A")) {
				missingArgs(sheetInfo, data, stepNumber);
				return 0;
			}

			locatorValue = getDataParam(sheetInfo, data, variables, data.get(4));
			dataParam = getDataParam(sheetInfo, data, variables, data.get(5));
			
			query = data.get(4).toString();
			appLogs.info("Executing query -  " + query);

			/* Creating a new connection */
			Connection con = (Connection) DriverManager.getConnection(URL,
					userName, password);

			/* Creating a new statement object */
			Statement stmt = (Statement) con.createStatement();

			/* Execute the SQL Query to generate the result set */
			ResultSet resultSet = stmt.executeQuery(query);

			while (resultSet.next()) {
				if (!(resultSet.getString(1).equals(dataParam))) {
					appLogs.error("Execution Failed: " + dataParam
							+ " is not present in the data base in "
							+ stepNumber + " of Sheet " + sheetInfo[1]);
					// main.resultDataRow.add("Execution Failed at Step " +
					// stepNumber + " due to missing arguments.");
					// main.resultDataRow.add("N/A");
					return 0;
				}

				else {
					appLogs.info("Data Matched successfully");
				}

			}

			return 1;
		}

		catch (SQLException e) {
			appLogs.error("Execution Failed :  For Action " + data.get(2)
					+ " in Step number " + stepNumber + " of Sheet "
					+ sheetInfo[1]);
			appLogs.error(e.getMessage());
			getScreenshot(sheetInfo, stepNumber, e.getMessage()
					.substring(0, 50) + "...");

			return 0;
		} catch (Exception e) {
			String exceptionMessage = exceptionalCondition(sheetInfo, data,
					stepNumber, e);
			getScreenshot(sheetInfo, stepNumber, exceptionMessage);
			return 0;
		}

	}
	
	
	/**
	 * Quits the browser 
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param stepNumber - Currently executing row number 
	 * @param status - Status of the currently executing sheet for which quit is called.(0-FAIL , 1- PASS)
	 * @return - N/A
	 */
	public void quit(String [] sheetInfo, int status) {

		if (status == 1) {
			appLogs.info("< ------ Successfully Completed the execution for sheet : "+ sheetInfo[1] + " in File : " + sheetInfo[0] + "------ >");
			//createResultSet(sheetInfo[0], sheetName, "PASS", "N/A", "N/A");
			createResultSet(sheetInfo, "PASS", "N/A", "N/A");
		} else
			appLogs.error("< ------ Executon failed for sheet :" + sheetInfo[1] + " ------ >");

		driver.close();
		//driver.quit();
	}
	
	
	
	
	
	/* <------------------------------------------ Non-Selenium Methods ---------------------------------------------------- >*/

	/**
	 * Takes the data from a 'Locator Value' and 'Test Data/Options' column and adds and updates them for 
	 * any 'variable' provided in the 'variables' hashmap
	 * 
	 * @param variables - HashMap of the all the variables defined in the config sheet of the file
	 * @param data - contains the String provided in 'Locator Value' or 'Test Data/Options' column of a row.
	 * @return - Returns the updated string value
	 */	
	public String getDataParam(String[] sheetInfo, ArrayList<String> data,
			HashMap<String, String> variables, String param) {
		
		String commonSheetKey = sheetInfo[0]+":"+sheetInfo[1]+":"+data.get(0)+":"+data.get(7);
		appLogs.debug("Common Sheet keys is " + commonSheetKey);
		

		if(  allCommonSheetVars.get(commonSheetKey)!=null){
			for(String key: allCommonSheetVars.get(commonSheetKey).keySet()){
				
				if (param.indexOf(key) != -1){
					param= param.replace(key, allCommonSheetVars.get(commonSheetKey).get(key));	
					   appLogs.debug("Acutal Data parameter is in allCommonSheetVars Map " + param);
    				    return param.trim();
				}				
			}
		}
			
		for (String key : variables.keySet()) {
			appLogs.debug("Finding data parameter for : " + key + " : "+ variables.get(key));

			if (param.indexOf(key) != -1){
				param= param.replace(key, variables.get(key));
			    appLogs.debug("Acutal Data parameter is in Variables " + param);
			    return param.trim();	
			}
				
		}
		
		appLogs.debug("Acutal Data parameter is " + param);
		return param.trim();

	}

	/**
	 * Adds error message, updates final Result Set in case of any missing arguments in a row.
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - Cell data from the currently executing row 
	 * @param stepNumber - Currently executing row number 
	 */	
	public void missingArgs(String[] sheetInfo, ArrayList data, String stepNumber) {
		appLogs.debug("Inside the missingArgs Method");
		
		appLogs.error("Execution Failed: Missing Arugments for the Action " + data.get(2) 
				    + " in Step number " + stepNumber + " of Sheet "+ sheetInfo[1]);
		
		createResultSet(sheetInfo, "FAIL","Missing Arugments for the Action " + data.get(2)
						+ " in Step number " + stepNumber + " of Sheet "+ sheetInfo[1], "N/A");
		
		quit(sheetInfo, 0); //Quit the browser due to Execution Failure
	}
	
	
	/**
	 * Adds error message, updates final Result Set in case of any incorrect Locator Type in a row.
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - Cell data from the currently executing row 
	 * @param stepNumber - Currently executing row number 
	 */	
	public void incorrectLocatorType(String[] sheetInfo, ArrayList data, String stepNumber){
		appLogs.debug("Inside the incorrectLocatorType Method");
		
		appLogs.error("Execution Failed: Incorrect Locator Type for the Action "+ data.get(2) 
				    + " in Step number " + stepNumber + " of Sheet "+ sheetInfo[1]);
		
		createResultSet(sheetInfo, "FAIL","Incorrect Locator Type for the Action " + data.get(2)
						+ " in Step number " + stepNumber + " of Sheet "+ sheetInfo[1], "N/A");
		
		quit(sheetInfo, 0);  //Quit the browser due to Execution Failure
		
	}

	/**
	 * Adds error message, updates final Result Set in case of incorrect variable declaration in the 
	 * Test Data/Options Column
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - Cell data from the currently executing row 
	 * @param stepNumber - Currently executing row number 
	 * @param currentSheet - Currently Executing Sheet
	 */	
	public void incorrectVariableDeclaration(String [] sheetInfo, ArrayList data, String stepNumber){
		
        appLogs.debug("Inside the incorrectVariableAssgnment Method");
		
		appLogs.error("Execution Failed: Incorrect Variable Declaration for the Action "+ data.get(2) 
				    + " in Step number " + stepNumber + " of Sheet "+ sheetInfo[1]  + "USE format -  var <variable name>");
		
		createResultSet(sheetInfo, "FAIL","Incorrect Variable Declaration for the Action " + data.get(2)
						+ " in Step number " + stepNumber + " of Sheet "+ sheetInfo[1] + " USE format - var <variable name>", "N/A");
		
		quit(sheetInfo, 0);  //Quit the browser due to Execution Failure
		
	}
	
	/**
	 * Logs error message Set in case of any exceptional condition. 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param data - Cell data from the currently executing Sheet
	 * @param stepNumber - Currently executing row number 
	 * @param e - Exception class object
	 * @return - returns to the Exception message
	 */

	public String exceptionalCondition(String [] sheetInfo, ArrayList<String> data,String stepNumber, Exception e) {
		appLogs.debug("Inside the exceptionalCondition method");
		appLogs.error("Exception Raised :  For Action " + data.get(2)+ " in Step number " + stepNumber + " of Sheet" + sheetInfo[1]);
		
		appLogs.error(e.getMessage());
		return e.getMessage().toString();
	}
	
	
	/**
	 * Takes screenshot in case of any failure
	 * 
	 * @param sheetInfo -  String array containing info of the sheet config- file name, current sheet name
	 * browser,OS if any, Node URL if any.
	 * @param stepNumber - Currently executing row number 
	 * @param exceptionMessage - Exception message that was raised in the current stepNumber
	 * @return - N/A
	 */
	public void getScreenshot(String [] sheetInfo,String stepNumber, String exceptionMessage) {
          
		String screenshotName = "";
		if(!enable_grid){
		  screenshotName = sheetInfo[0] + "-" + sheetInfo[1] + "-" + "Step"+ stepNumber;
		}
		else{
			screenshotName = sheetInfo[0] + "-" + sheetInfo[1] + "-" + sheetInfo[2] + "-"+ sheetInfo[4] + "-"+ "Step"+ stepNumber;
		}
		
		File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		try {
			FileUtils.copyFile(scrFile, new File(System.getProperty("user.dir")+ "/Results/Screenshots/" + screenshotName + ".png"));
		} 
		catch (IOException e) {// TODO Auto-generated catch block
			e.printStackTrace();
		}

		createResultSet(sheetInfo, "FAIL", exceptionMessage,screenshotName);
		quit(sheetInfo, 0);
	}
	
	public void testMethod(){
		appLogs.info("This is test" + Thread.currentThread().getName());
	}

}

