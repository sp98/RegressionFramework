package reporting;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;

public class HTMLReport {

	public static FileWriter htmlFile = null;
	public static BufferedWriter out = null;
	

	/**
	 * Creates a new Results Directory, if not present.
	 * Deletes all the existing files form the Results Directory if already present.
	 * Creates a new Screenshot Directory inside the Reullts Directory
	 * 
	 */
	public static void checkResultDir() {
		
		File dir = new File(System.getProperty("user.dir") + "/Results");
		/*Delete all the previous files and folders in the Reuslts Directory if its not empty */
		if (dir.exists()) {
			File[] allFiles = dir.listFiles();

			for (int i = 0; i < allFiles.length; i++) {
				allFiles[i].delete();
			}

		} else {
			new File(System.getProperty("user.dir") + "/Results").mkdir();

		}

		/* Make a new screenshot directory inside the Results Directory */
		new File(System.getProperty("user.dir") + "/Results/Screenshots").mkdir();
	
	}

	
	/**
	 * Creates a HTML report after all the files have been executed.
	 * 
	 * @param resultSet - Final Result Set after complete execution of all the sheets.
	 * @param executableFiles - Total Number of Executable Files. 
	 */
	
	 public static void makeHtmlReport( LinkedHashMap<String, ArrayList<String>> resultSet, ArrayList<String> executableFiles){ 
		 
		 htmlFile = null; 
		 out = null;
		 
		    DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
			DateFormat timeFormat = new SimpleDateFormat("HH:mm:ss");
			Date date = new Date();	
	  
	     try { 
		   htmlFile = new FileWriter(System.getProperty("user.dir") +"/Results/index.html", true); 
		   out = new BufferedWriter(htmlFile);
	       out.newLine(); 
	       } 
	     
	    catch (IOException e){ 
	       e.printStackTrace();
	      }
	  
	    
	     try {
			out.write("	<html>\n"
					+ "		<head>\n"
					+ "		<title>Test Execution Report</title>\n"
					+ "		</head>\n"
					+ "			<body>\n"
					+ "				<br>\n"
					+ "				<h1 style='text-align:center;'> FLOW TEST REPORT: Generated On "
					+ dateFormat.format(date) + " at "
					+ timeFormat.format(date) + "</h1>\n" + "				<br>\n"
					+ "				<br>\n" + "			</body>\n" + "		</html>\n");

		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
	  
	 
	 /* For loop to print the result data for each executable file */ 
	  for (String fileName : executableFiles) {
		  
		  try { 	  
			  out.write("	<html>\n"
		          + "		<head>\n" 
				  + "		<title>Test Execution Report</title>\n" 
		          + "		</head>\n" 
		          + "			<body>\n" + "				<br>\n"
		          + "				<table style=' width:20%; border: 1px solid gray;  margin-left: 100px; '>\n"
				  + "            <br>\n"
				  + "            <br>\n"
				  + "            <tr>\n"
				  + "				<td style='color:Black; border: hidden hidden solid hidden 1px black; align:center; background-color: #d9edf7'><b> FILE NAME </b></td>\n"
				  + "						<td style='color:Black; border: hidden hidden solid hidden 1px black 1px black; align:center; background-color:	#F8F8F8  '><b>"
				  +                     fileName + "</b></td>\n" + "</tr>\n" + "</table>\n " 
				  + "              <table style=' width:80%; border: 1px solid black; align:center;align:center;margin-left: 100px '>\n"
				  + "              	<thead>\n" + "						"
				  + "                  <tr>\n" 
				  + "						<th style=' width:10%; border: solid 1px black; background-color: Crimson;"
				  + "							color: white;'> SHEET NAME </th>\n" +
				  " 							<th style='width:5%; border: solid 1px black; background-color: Crimson;"
				  + " 							color: white;'> STATUS </th>\n" +
				  "							<th style=' width:55%; border: solid 1px black; background-color: Crimson;"
				  + "							color: white;'> ERROR MESSAGE </th>\n" +
				  "							<th style='width:20%; border: solid 1px black; background-color: Crimson;"
				  + "							color: white;'> SCREENSHOT NAME </th>\n" + "						"
				  + "                  </tr>\n" 
				  + "                  </thead>\n" 
				  + "					<tbody>\n");
			  
		         } 
				                                    
				  
				  catch (IOException e1) {   
					  e1.printStackTrace();
					  }
		  
		  /* for loop for printing result data for each sheet inside the file */
		  for (String keys : resultSet.keySet()){
			  
			  if (keys.split(":")[0].equals(fileName)){
				  
				  String sheetName =keys.split(":")[1];
				  
				  try { 
					  
					  out.write( "<td style='color:black; border: solid 1px black; align:center;'> "+ sheetName + "</td>\n");
					  
					  if(resultSet.get(keys).get(0).equalsIgnoreCase("PASS"))  
						  out.write("<td style='color:black; border: solid 1px black; align:center;' bgcolor=#66FF66;>"
				                    + resultSet.get(keys).get(0) + "</td>\n"); 
					     else  
					    	 out.write( "<td style='color:black; border: solid 1px black; align:center;' bgcolor=#FF6666;>"
				                    + resultSet.get(keys).get(0) + "</td>\n"); 
				  
				      out.write( "<td style='color:black; border: solid 1px black; align:center;'>" + resultSet.get(keys).get(1) + "</td>\n"
				                   +"            <td style='color:black; border: solid 1px black; align:center;'>"
				  + resultSet.get(keys).get(2) + "</td>\n" + "	</tr>\n");	
				  
				  } 
				  
				  catch (IOException e) { 
				    e.printStackTrace(); }
				  
				  } //if block ends here
			  	  
		  } // for loop print result data ends here
		  
			  }
	    
	   try { out.write("<br>\n");
		     out.write("	</tbody>\n" + "</table>\n" +
			  "                         <br><br>\n" + "</html>\n"); 
			  } 
			  
			  catch
			  (IOException e1) { // TODO Auto-generated catch block
			  e1.printStackTrace(); }
	      
	  try {  
		  out.close(); 
		  } 
	  
	   catch (IOException e) { // TODO Auto-generated catch block
		  e.printStackTrace(); }
	  
		  }	  
		  
	  }
	  